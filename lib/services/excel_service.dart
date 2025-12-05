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

  static Future<String> exportGuests(List<Guest> guests) async {
    // Request storage permission
    var status = await Permission.manageExternalStorage.request();
    if (!status.isGranted) {
      // If manage external storage is not granted, try with storage permission
      status = await Permission.storage.request();
      if (!status.isGranted) {
        throw Exception("Storage permission denied. Please grant storage permission to save the file.");
      }
    }

    // Create Excel file
    final excel = excel_pkg.Excel.createExcel();
    final sheet = excel['Verified Guests'];
    
    // Add header row
    sheet.appendRow([
      excel_pkg.TextCellValue('Name'),
      excel_pkg.TextCellValue('Phone'),
      excel_pkg.TextCellValue('Verification Time'),
      excel_pkg.TextCellValue('Status')
    ]);
    
    // Add data rows
    final now = DateTime.now().toIso8601String();
    for (var guest in guests) {
      sheet.appendRow([
        excel_pkg.TextCellValue(guest.name),
        excel_pkg.TextCellValue(guest.phone),
        excel_pkg.TextCellValue(now),
        excel_pkg.TextCellValue('Verified')
      ]);
    }
    
    try {
      Directory? directory;
      
      if (Platform.isAndroid) {
        // For Android, try to get the public Downloads directory
        directory = Directory('/storage/emulated/0/Download');
        
        // If the directory doesn't exist, try to create it
        if (!await directory.exists()) {
          await directory.create(recursive: true);
        }
        
        // Check if we can write to the directory
        final testFile = File('${directory.path}/.test');
        try {
          await testFile.writeAsString('test');
          await testFile.delete();
        } catch (e) {
          // If we can't write to the directory, fall back to the app's documents directory
          directory = await getApplicationDocumentsDirectory();
        }
      } else if (Platform.isIOS) {
        // For iOS, use the documents directory
        directory = await getApplicationDocumentsDirectory();
      } else {
        // For other platforms, use the documents directory
        directory = await getApplicationDocumentsDirectory();
      }
      
      // Create a unique filename with timestamp
      final fileName = 'Bandhan_Verified_Guests_${DateTime.now().millisecondsSinceEpoch}.xlsx';
      final file = File('${directory.path}/$fileName');
      
      // Save the file
      await file.writeAsBytes(excel.encode()!);
      
      // If we're on Android and the file was saved to the app's directory,
      // try to copy it to the public Downloads folder
      if (Platform.isAndroid && !directory.path.contains('Download')) {
        try {
          final downloadsDir = Directory('/storage/emulated/0/Download');
          if (!await downloadsDir.exists()) {
            await downloadsDir.create(recursive: true);
          }
          
          final publicFile = File('${downloadsDir.path}/$fileName');
          await file.copy(publicFile.path);
          await OpenFile.open(publicFile.path);
          return publicFile.path;
        } catch (e) {
          print('Could not save to public Downloads folder: $e');
          // Continue to return the original file path
        }
      }
      
      // Open the file
      await OpenFile.open(file.path);
      
      return file.path;
    } catch (e) {
      throw Exception('Failed to save file: $e');
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
    List<Guest> guests, {
    String? fileName,
  }) async {
    if (guests.isEmpty) {
      throw const FormatException('No verified guests to export.');
    }

    final excel = excel_pkg.Excel.createExcel();
    final sheet = excel['Verified Guests'];
    sheet.appendRow([
      excel_pkg.TextCellValue('Name'),
      excel_pkg.TextCellValue('Phone'),
    ]);

    for (final guest in guests) {
      sheet.appendRow([
        excel_pkg.TextCellValue(guest.name),
        excel_pkg.TextCellValue(guest.phone),
      ]);
    }

    final bytes = excel.encode();
    if (bytes == null) {
      throw const FormatException('Unable to encode Excel file.');
    }

    final directory = await getApplicationDocumentsDirectory();
    final ts = DateTime.now().millisecondsSinceEpoch;
    final safeFileName = fileName?.trim().isNotEmpty == true
        ? fileName!.trim()
        : 'verified_guests_$ts';
    final file = File('${directory.path}/$safeFileName.xlsx');
    await file.writeAsBytes(bytes, flush: true);
    return file;
  }
}

