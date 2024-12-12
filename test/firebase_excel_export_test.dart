import 'package:flutter_test/flutter_test.dart';
import 'package:firebase_excel_export/firebase_excel_export.dart';

void main() {
  test('Exports data to Excel without errors', () async {
    expect(
      () async => await FirebaseExcelExporter.exportToExcel(
        collectionName: 'test_collection',
      ),
      returnsNormally,
    );
  });
}
