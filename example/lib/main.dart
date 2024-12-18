import 'package:flutter/material.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:firebase_excel_export/firebase_excel_export.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  // Initialize Firebase
  await Firebase.initializeApp();

  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(title: const Text('Firestore to Excel Example')),
        body: Center(
          child: ElevatedButton(
            onPressed: exportFilteredUsers,
            child: const Text('Export Active Users'),
          ),
        ),
      ),
    );
  }

  /// Exports active users from Firestore to an Excel file.
  Future<void> exportFilteredUsers() async {
    try {
      await FirebaseExcelExporter.exportToExcel(
        collectionName: 'users',
        headers: ['Email', 'Name', 'Status'], // Custom headers for the Excel file
        mapper: (data) => [
          data['email'],       // Map Firestore fields to headers
          data['name'],
          data['status'],
        ],
        queryBuilder: (query) => query.where('status', isEqualTo: 'active'), // Firestore filter
        fileName: 'active_users_export.xlsx', // Output file name
      );

    } catch (e) {
      debugPrint(e.toString());
    }
  }
}
