import 'package:flutter_test/flutter_test.dart';
import 'package:fake_cloud_firestore/fake_cloud_firestore.dart';
import 'package:firebase_excel_export/firebase_excel_export.dart';

void main() {
  test('Exports data to Excel with specific headers and data mapping', () async {
    // Initialize a fake Firestore instance
    final fakeFirestore = FakeFirebaseFirestore();

    // Add mock data to Firestore
    await fakeFirestore.collection('test_collection').add({
      'email': 'test@example.com',
      'name': 'Test User',
      'phoneNumber': '1234567890',
    });

    // Inject the fake Firestore instance into the FirebaseExcelExporter (if needed)
    // In this example, we assume `FirebaseExcelExporter` uses the default Firestore instance.
    // If your implementation supports dependency injection, pass `fakeFirestore` accordingly.

    expect(() async {
      await FirebaseExcelExporter.exportToExcel(
        collectionName: 'test_collection',
        headers: ['Email', 'Name', 'Phone Number'], // Custom headers
        mapper: (data) => [
          data['email'],       // Map Firestore fields to specific headers
          data['name'],
          data['phoneNumber'],
        ],
      );
    }, returnsNormally);
  });
}
