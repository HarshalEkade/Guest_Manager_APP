class Guest {
  const Guest({
    required this.name,
    required this.phone,
    this.count = 0,
  });

  final String name;
  final String phone;
  final int count;

  factory Guest.fromExcelRow(List<dynamic> row) {
    String valueAt(int index) {
      if (index >= row.length) return '';
      final cell = row[index];
      if (cell == null) return '';
      return cell.toString().trim();
    }

    final firstCell = valueAt(0);
    final secondCell = valueAt(1);
    final thirdCell = valueAt(2);
    final fourthCell = valueAt(3); // Count column

    final hasSerialNumber =
        RegExp(r'^\d+$').hasMatch(firstCell.replaceAll(' ', ''));

    final name = hasSerialNumber && secondCell.isNotEmpty ? secondCell : firstCell;
    final phone = hasSerialNumber && secondCell.isNotEmpty ? thirdCell : secondCell;
    final countStr = hasSerialNumber && fourthCell.isNotEmpty ? fourthCell :
                     (!hasSerialNumber && thirdCell.isNotEmpty) ? thirdCell : '';

    int count = 0;
    if (countStr.isNotEmpty) {
      final parsedCount = int.tryParse(countStr);
      if (parsedCount != null) {
        count = parsedCount;
      }
    }

    return Guest(
      name: name,
      phone: phone,
      count: count,
    );
  }

  Map<String, dynamic> toMap() => {
        'Name': name,
        'Phone': phone,
        'Count': count,
      };
}

