import 'dart:io';
import 'package:excel/excel.dart' as excel_pkg;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:open_file/open_file.dart';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

import '../models/guest.dart';

class ExcelService {
  Future<List<Guest>> loadGuestList(String filePath) async {
    final file = File(filePath);
    if (!await file.exists()) {
      throw const FileSystemException('Guest list file not found.');
    }

    final bytes = await file.readAsBytes();

    try {
      return _parseWithExcelPackage(bytes);
    } catch (error) {
      return _parseWithDecoder(bytes, error);
    }
  }

  List<Guest> _parseWithExcelPackage(List<int> bytes) {
    final excel = excel_pkg.Excel.decodeBytes(bytes);
    if (excel.tables.isEmpty) return [];

    final firstTable = excel.tables.values.first;
    final normalizedRows = firstTable.rows
        .map((row) => row.map((cell) => cell?.value).toList())
        .toList();
    return _rowsToGuests(normalizedRows);
  }

  List<Guest> _parseWithDecoder(List<int> bytes, Object originalError) {
    try {
      final decoder = SpreadsheetDecoder.decodeBytes(bytes);
      if (decoder.tables.isEmpty) return [];
      final table = decoder.tables.values.first;
      return _rowsToGuests(table.rows);
    } catch (fallbackError) {
      throw FormatException(
        'Failed to load file. Original error: $originalError. '
        'Decoder error: $fallbackError',
      );
    }
  }

  List<Guest> _rowsToGuests(List<List<dynamic>> rows) {
    if (rows.isEmpty) return [];
    final dataRows = rows.skip(1);
    final guests = <Guest>[];

    for (final row in dataRows) {
      final guest = Guest.fromExcelRow(row);
      if (guest.name.isEmpty && guest.phone.isEmpty) continue;
      guests.add(guest);
    }
    return guests;
  }

  Future<File> exportVerifiedGuests(
    List<Guest> allGuests,
    List<Guest> verifiedGuests,
    Map<String, int> verificationCounts,
    String originalFilePath, {
    String? fileName,
  }) async {
    if (verifiedGuests.isEmpty) {
      throw const FormatException('No verified guests to export.');
    }

    // 1. Update original Excel file with verification counts
    await _updateOriginalExcelFile(
      originalFilePath,
      allGuests,
      verificationCounts,
    );

    // 2. Create the new export file with ONLY verified guests and their counts
    final excel = excel_pkg.Excel.createExcel();
    final sheet = excel['Verified Guests'];
    sheet.appendRow([
      excel_pkg.TextCellValue('Name'),
      excel_pkg.TextCellValue('Phone'),
      excel_pkg.TextCellValue('Count'),
    ]);

    // Add only verified guests with their verification counts to the new file
    for (final guest in verifiedGuests) {
      final normalizedPhone = _normalizePhone(guest.phone);
      final count = verificationCounts[normalizedPhone] ?? 0;
      sheet.appendRow([
        excel_pkg.TextCellValue(guest.name),
        excel_pkg.TextCellValue(guest.phone),
        excel_pkg.TextCellValue(count.toString()),
      ]);
    }

    final bytes = excel.encode();
    if (bytes == null) {
      throw const FormatException('Unable to encode Excel file.');
    }

    final directory = await getApplicationDocumentsDirectory();
    final timestamp = DateTime.now().millisecondsSinceEpoch;
    final safeFileName = fileName?.trim().isNotEmpty == true
        ? fileName!.trim()
        : 'verified_guests_$timestamp';
    final file = File('${directory.path}/$safeFileName.xlsx');
    await file.writeAsBytes(bytes, flush: true);
    return file;
  }

  Future<void> _updateOriginalExcelFile(
    String originalFilePath,
    List<Guest> allGuests,
    Map<String, int> verificationCounts,
  ) async {
    final originalFile = File(originalFilePath);
    if (!await originalFile.exists()) return;

    // Load the original Excel file
    final bytes = await originalFile.readAsBytes();
    final originalExcel = excel_pkg.Excel.decodeBytes(bytes);

    if (originalExcel.tables.isEmpty) return;

    final sheet = originalExcel.tables.values.first;

    // Check if Count column already exists (column D)
    bool countColumnExists = false;
    if (sheet.maxCols >= 4) {
      // Check if column D header is 'Count'
      final headerCell = sheet.rows.isEmpty ? null : sheet.rows.first.length > 3 ? sheet.rows.first[3] : null;
      if (headerCell != null && headerCell.toString().trim() == 'Count') {
        countColumnExists = true;
      }
    }

    if (!countColumnExists) {
      // Add Count column header to column D
      if (sheet.rows.isNotEmpty) {
        sheet.rows[0].add(excel_pkg.TextCellValue('Count'));
      }
    }

    // Update the rows with verification counts
    for (int i = 1; i < sheet.rows.length && i <= allGuests.length; i++) {
      final row = sheet.rows[i];
      final guest = allGuests[i - 1]; // Match by position

      final normalizedPhone = _normalizePhone(guest.phone);
      final count = verificationCounts[normalizedPhone] ?? 0;

      if (row.length > 3) {
        // Update existing count column
        row[3] = excel_pkg.TextCellValue(count.toString());
      } else {
        // Add new count column
        row.add(excel_pkg.TextCellValue(count.toString()));
      }
    }

    // Save the updated Excel file
    final updatedBytes = originalExcel.encode();
    if (updatedBytes != null) {
      await originalFile.writeAsBytes(updatedBytes);
    }
  }

  String _normalizePhone(String value) =>
      value.replaceAll(RegExp(r'[^0-9+]'), '').replaceFirst(RegExp(r'^\+'), '');
}

