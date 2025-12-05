class Guest {
  const Guest({
    required this.name,
    required this.phone,
  });

  final String name;
  final String phone;

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

    final hasSerialNumber =
        RegExp(r'^\d+$').hasMatch(firstCell.replaceAll(' ', ''));

    final name = hasSerialNumber && secondCell.isNotEmpty ? secondCell : firstCell;
    final phone =
        hasSerialNumber && secondCell.isNotEmpty ? thirdCell : secondCell;

    return Guest(
      name: name,
      phone: phone,
    );
  }

  Map<String, dynamic> toMap() => {
        'Name': name,
        'Phone': phone,
      };
}

