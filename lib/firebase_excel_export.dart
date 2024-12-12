import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel/excel.dart';
import 'package:universal_html/html.dart' as html;

class FirebaseExcelExporter {
  /// Exports a Firestore collection to an Excel file.
  ///
  /// [collectionName] - The name of the Firestore collection to fetch data from.
  /// [headers] - The list of headers for the Excel file.
  /// [mapper] - A function that maps Firestore document data to a list of values for each row.
  /// [fileName] - The name of the resulting Excel file (default: 'export.xlsx').
  /// [queryBuilder] - An optional function to customize the Firestore query.
  static Future<void> exportToExcel({
    required String collectionName,
    required List<String> headers,
    required List<dynamic> Function(Map<String, dynamic>) mapper,
    String fileName = 'export.xlsx',
    Query Function(Query query)? queryBuilder,
  }) async {
    try {
      // Build Firestore query
      Query query = FirebaseFirestore.instance.collection(collectionName);
      if (queryBuilder != null) {
        query = queryBuilder(query);
      }

      // Fetch data from Firestore
      QuerySnapshot snapshot = await query.get();

      if (snapshot.docs.isEmpty) {
        throw Exception(
            'No data found in the Firestore collection: $collectionName');
      }

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
        final row = mapper(data);

        for (int i = 0; i < row.length; i++) {
          sheetObject
              .cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: rowIndex))
              .value = row[i] != null
              ? TextCellValue(row[i].toString())
              : TextCellValue('');
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
