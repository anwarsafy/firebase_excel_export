library firebase_excel_export;

import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel/excel.dart';
import 'package:universal_html/html.dart' as html;

class FirebaseExcelExporter {
  /// Exports a Firestore collection to an Excel file.
  ///
  /// [collectionName] - The name of the Firestore collection to fetch data from.
  /// [fileName] - The name of the resulting Excel file (default: 'export.xlsx').
  static Future<void> exportToExcel({
    required String collectionName,
    String fileName = 'export.xlsx',
  }) async {
    try {
      // Fetch data from Firestore
      QuerySnapshot snapshot =
          await FirebaseFirestore.instance.collection(collectionName).get();

      if (snapshot.docs.isEmpty) {
        throw Exception(
            'No data found in the Firestore collection: $collectionName');
      }

      // Determine dynamic headers from the first document
      final firstDoc = snapshot.docs.first.data() as Map<String, dynamic>;
      final headers = firstDoc.keys.toList();

      // Initialize Excel
      var excel = Excel.createExcel();
      Sheet sheetObject = excel['Sheet1'];

      // Write headers to the first row
      for (int i = 0; i < headers.length; i++) {
        sheetObject
            .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0))
            .value = TextCellValue(headers[i]);
      }

      // Write data rows
      int rowIndex = 1;
      for (var doc in snapshot.docs) {
        final data = doc.data() as Map<String, dynamic>;

        for (int i = 0; i < headers.length; i++) {
          final value = data[headers[i]];
          sheetObject
                  .cell(CellIndex.indexByColumnRow(
                      columnIndex: i, rowIndex: rowIndex))
                  .value =
              value != null
                  ? TextCellValue(value.toString())
                  : TextCellValue(''); // Changed here
        }
        rowIndex++;
      }

      // Generate Excel file
      final bytes = excel.save();
      final blob = html.Blob([bytes],
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

      // Create a download link and trigger the download
      final url = html.Url.createObjectUrlFromBlob(blob);
      html.AnchorElement(href: url)
        ..setAttribute('download', fileName)
        ..click();

      // Revoke the Blob URL to release resources
      html.Url.revokeObjectUrl(url);
    } catch (e) {
      throw Exception('Error exporting to Excel: $e');
    }
  }
}
