import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

import 'models/guest.dart';
import 'screens/qr_scanner_page.dart';
import 'services/excel_service.dart';

void main() {
  runApp(const GuestVerificationApp());
}

class GuestVerificationApp extends StatelessWidget {
  const GuestVerificationApp({super.key});

  @override
  Widget build(BuildContext context) {
    return LayoutBuilder(
      builder: (context, constraints) {
        // Set a base screen size for responsive design
        final screenWidth = constraints.maxWidth;
        final screenHeight = constraints.maxHeight;
        final isSmallScreen = screenWidth < 600;
        
        return MaterialApp(
          title: 'Bandhan Logger',
          debugShowCheckedModeBanner: false, // Remove debug banner
          theme: ThemeData(
            colorScheme: ColorScheme.fromSeed(
              seedColor: const Color(0xFF5D54A4),
            ),
            textTheme: TextTheme(
              headlineSmall: TextStyle(
                fontWeight: FontWeight.w600,
                fontSize: isSmallScreen ? 18 : 22,
              ),
              titleMedium: TextStyle(
                fontWeight: FontWeight.w500,
                fontSize: isSmallScreen ? 14 : 16,
              ),
              bodyLarge: TextStyle(
                fontSize: isSmallScreen ? 14 : 16,
              ),
            ),
        appBarTheme: const AppBarTheme(
          centerTitle: false,
          elevation: 0,
          titleTextStyle: TextStyle(
            fontSize: 20,
            fontWeight: FontWeight.w600,
          ),
        ),
        cardTheme: CardTheme(
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
          elevation: 0,
          color: Colors.white,
        ),
        useMaterial3: true,
      ),
      home: const GuestVerificationPage(),
    );
  }
    );
}
}


class GuestVerificationPage extends StatefulWidget {
  const GuestVerificationPage({super.key});

  @override
  State<GuestVerificationPage> createState() => _GuestVerificationPageState();
}

class _GuestVerificationPageState extends State<GuestVerificationPage> {
  final ExcelService _excelService = ExcelService();
  final TextEditingController _manualEntryController = TextEditingController();

  List<Guest> _guests = [];
  List<Guest> _verifiedGuests = [];
  bool _isLoading = false;
  String? _activeGuestFile;
  Guest? _lastVerified;
  String? _statusMessage;

  @override
  void dispose() {
    _manualEntryController.dispose();
    super.dispose();
  }

  Future<void> _pickGuestFile() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );

      if (result == null || result.files.single.path == null) return;
      setState(() {
        _isLoading = true;
        _statusMessage = null;
      });

      final path = result.files.single.path!;
      final guests = await _excelService.loadGuestList(path);

      setState(() {
        _guests = guests;
        _activeGuestFile = result.files.single.name;
        _isLoading = false;
        _verifiedGuests = [];
        _lastVerified = null;
        _statusMessage = 'Loaded ${guests.length} guest records.';
      });
    } catch (error) {
      setState(() {
        _isLoading = false;
        _statusMessage = 'Failed to load file: $error';
      });
    }
  }

  Future<void> _scanQrCode() async {
    if (_guests.isEmpty) {
      _showSnackBar('Upload the guest list first.');
      return;
    }
    final qrValue = await Navigator.push<String>(
      context,
      MaterialPageRoute(builder: (_) => const QrScannerPage()),
    );
    if (qrValue == null || qrValue.isEmpty) return;
    _verifyGuest(qrValue.trim());
  }

  void _verifyGuest(String input) {
    if (_guests.isEmpty) {
      _showSnackBar('Upload the guest list first.');
      return;
    }

    final cleanedInput = input.trim();
    if (cleanedInput.isEmpty) {
      _showSnackBar('Invalid QR/phone value.');
      return;
    }

    final normalizedInput = _normalizePhone(cleanedInput);

    Guest? match;
    for (final guest in _guests) {
      if (_normalizePhone(guest.phone) == normalizedInput) {
        match = guest;
        break;
      }
    }

    if (match == null) {
      _showSnackBar('Guest not found.');
      setState(() => _statusMessage = 'No guest mapped to $cleanedInput');
      return;
    }

    final alreadyVerified = _verifiedGuests.any(
      (g) => _normalizePhone(g.phone) == normalizedInput,
    );
    if (alreadyVerified) {
      _showSnackBar('${match.name} already verified.');
      return;
    }

    setState(() {
      _verifiedGuests = [..._verifiedGuests, match!];
      _lastVerified = match;
      _statusMessage = 'Verified ${match.name}';
    });
    _showSnackBar('Guest verified: ${match.name}');
  }

  Future<void> _exportVerified() async {
    if (_verifiedGuests.isEmpty) {
      _showSnackBar('No verified guests to export.');
      return;
    }

    try {
      setState(() => _isLoading = true);
      final file = await _excelService.exportVerifiedGuests(_verifiedGuests);
      if (!mounted) return;
      setState(() => _isLoading = false);
      _showSnackBar('Verified list saved at ${file.path}');
    } catch (error) {
      setState(() => _isLoading = false);
      _showSnackBar('Export failed: $error');
    }
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text(message)),
    );
  }

  String _normalizePhone(String value) =>
      value.replaceAll(RegExp(r'[^0-9+]'), '').replaceFirst(RegExp(r'^\+'), '');

  @override
  Widget build(BuildContext context) {
    final gradient = LinearGradient(
      colors: [
        Theme.of(context).colorScheme.primary.withOpacity(0.12),
        Theme.of(context).colorScheme.secondaryContainer.withOpacity(0.3),
      ],
      begin: Alignment.topLeft,
      end: Alignment.bottomRight,
    );

    return Scaffold(
      extendBodyBehindAppBar: true,
      appBar: AppBar(
        title: const Text('Event Guest Verification'),
        backgroundColor: Colors.transparent,
        actions: [
          IconButton(
            icon: const Icon(Icons.cloud_download_outlined),
            tooltip: 'Download verified list',
            onPressed: _exportVerified,
          ),
        ],
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: _isLoading ? null : _scanQrCode,
        icon: const Icon(Icons.qr_code_scanner),
        label: const Text('Scan QR'),
      ),
      body: Container(
        decoration: BoxDecoration(gradient: gradient),
        child: SafeArea(
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              children: [
                _SummaryHeader(
                  totalGuests: _guests.length,
                  verifiedGuests: _verifiedGuests.length,
                  fileName: _activeGuestFile,
                  isLoading: _isLoading,
                  onUploadTap: _pickGuestFile,
                ),
                const SizedBox(height: 16),
                _ManualEntryCard(
                  controller: _manualEntryController,
                  onSubmit: (value) {
                    _verifyGuest(value);
                    _manualEntryController.clear();
                  },
                  statusMessage: _statusMessage,
                ),
                const SizedBox(height: 16),
                Expanded(
                  child: _VerifiedGuestsList(
                    guests: _verifiedGuests,
                    lastVerified: _lastVerified,
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _SummaryHeader extends StatelessWidget {
  const _SummaryHeader({
    required this.totalGuests,
    required this.verifiedGuests,
    required this.onUploadTap,
    this.fileName,
    this.isLoading = false,
  });

  final int totalGuests;
  final int verifiedGuests;
  final String? fileName;
  final bool isLoading;
  final VoidCallback onUploadTap;

  @override
  Widget build(BuildContext context) {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        'Guest list',
                        style: Theme.of(context).textTheme.titleMedium,
                      ),
                      const SizedBox(height: 4),
                      Text(
                        fileName ?? 'No file selected',
                        style: Theme.of(context)
                            .textTheme
                            .bodyMedium
                            ?.copyWith(color: Colors.grey[600]),
                      ),
                    ],
                  ),
                ),
                FilledButton.icon(
                  onPressed: isLoading ? null : onUploadTap,
                  icon: const Icon(Icons.upload_file),
                  label: Text(fileName == null ? 'Upload list' : 'Change'),
                ),
              ],
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                Expanded(
                  child: _StatCard(
                    label: 'Total guests',
                    value: totalGuests.toString(),
                    icon: Icons.groups_2_rounded,
                    color: Theme.of(context).colorScheme.primary,
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _StatCard(
                    label: 'Verified',
                    value: verifiedGuests.toString(),
                    icon: Icons.verified_rounded,
                    color: Colors.teal,
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}

class _StatCard extends StatelessWidget {
  const _StatCard({
    required this.label,
    required this.value,
    required this.icon,
    required this.color,
  });

  final String label;
  final String value;
  final IconData icon;
  final Color color;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(16),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(14),
        color: color.withOpacity(0.12),
      ),
      child: Row(
        children: [
          CircleAvatar(
            backgroundColor: color.withOpacity(0.18),
            child: Icon(icon, color: color),
          ),
          const SizedBox(width: 12),
          Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(
                value,
                style: Theme.of(context).textTheme.headlineSmall?.copyWith(
                      color: color,
                      fontWeight: FontWeight.bold,
                    ),
              ),
              Text(
                label,
                style: Theme.of(context).textTheme.bodyMedium,
              ),
            ],
          ),
        ],
      ),
    );
  }
}

class _ManualEntryCard extends StatelessWidget {
  const _ManualEntryCard({
    required this.controller,
    required this.onSubmit,
    this.statusMessage,
  });

  final TextEditingController controller;
  final void Function(String) onSubmit;
  final String? statusMessage;

  @override
  Widget build(BuildContext context) {
    final isError =
        statusMessage != null && statusMessage!.toLowerCase().contains('failed');

    return Card(
      child: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              'Manual verification',
              style: Theme.of(context).textTheme.titleMedium,
            ),
            const SizedBox(height: 8),
            TextField(
              controller: controller,
              keyboardType: TextInputType.phone,
              decoration: InputDecoration(
                hintText: 'Type phone number or QR content',
                prefixIcon: const Icon(Icons.phone_android),
                suffixIcon: IconButton(
                  icon: const Icon(Icons.check_circle),
                  onPressed: () => onSubmit(controller.text),
                ),
              ),
              onSubmitted: onSubmit,
            ),
            const SizedBox(height: 8),
            AnimatedSwitcher(
              duration: const Duration(milliseconds: 250),
              child: statusMessage == null
                  ? const SizedBox.shrink()
                  : Row(
                      key: ValueKey(statusMessage),
                      children: [
                        Icon(
                          isError ? Icons.error_outline : Icons.check_circle,
                          color: isError ? Colors.red : Colors.green,
                        ),
                        const SizedBox(width: 8),
                        Expanded(
                          child: Text(
                            statusMessage!,
                            style: TextStyle(
                              color: isError ? Colors.red : Colors.green,
                              fontWeight: FontWeight.w600,
                            ),
                          ),
                        ),
                      ],
                    ),
            ),
          ],
        ),
      ),
    );
  }
}

class _VerifiedGuestsList extends StatelessWidget {
  const _VerifiedGuestsList({
    required this.guests,
    this.lastVerified,
  });

  final List<Guest> guests;
  final Guest? lastVerified;

  @override
  Widget build(BuildContext context) {
    if (guests.isEmpty) {
      return Column(
        children: [
          const Spacer(),
          Icon(
            Icons.verified_outlined,
            size: 60,
            color: Theme.of(context).colorScheme.primary.withOpacity(0.5),
          ),
          const SizedBox(height: 12),
          Text(
            'No guests verified yet',
            style: Theme.of(context).textTheme.titleMedium,
          ),
          Text(
            'Scan a QR code or enter a phone number to begin.',
            textAlign: TextAlign.center,
            style: Theme.of(context).textTheme.bodyMedium,
          ),
          const Spacer(),
        ],
      );
    }

    return Card(
      child: Column(
        children: [
          ListTile(
            title: const Text('Verified guests'),
            subtitle: Text('${guests.length} guests checked in'),
            trailing: lastVerified == null
                ? null
                : Chip(
                    label: Text(lastVerified!.name.split(' ').first),
                    avatar: const Icon(Icons.flash_on, size: 16),
                  ),
          ),
          const Divider(height: 1),
          Expanded(
            child: ListView.separated(
              itemCount: guests.length,
              separatorBuilder: (_, __) => const Divider(height: 1),
              itemBuilder: (context, index) {
                final guest = guests[index];
                return ListTile(
                  leading: CircleAvatar(
                    backgroundColor:
                        Theme.of(context).colorScheme.primary.withOpacity(0.15),
                    child: Text(
                      '${index + 1}',
                      style: const TextStyle(fontWeight: FontWeight.bold),
                    ),
                  ),
                  title: Text(guest.name),
                  subtitle: Text(guest.phone),
                  trailing: const Icon(Icons.verified, color: Colors.green),
                );
              },
            ),
          ),
        ],
      ),
    );
  }
}
