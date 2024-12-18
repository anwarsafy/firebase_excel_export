import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel/excel.dart';
import 'package:universal_html/html.dart' as html;

/// A utility class for exporting Firestore collections to an Excel file.
///
/// This class allows customization of headers and data mapping, making it
/// flexible for various use cases.
///
/// Example usage:
/// ```dart
/// final exporter = FirebaseExcelExporter();
/// await exporter.exportCollection(
///   collectionPath: 'users',
///   filePath: '/path/to/save/file.xlsx',
///   headers: ['Name', 'Email', 'Age'],
///   dataMapper: (doc) => [doc['name'], doc['email'], doc['age']],
/// );
/// ```
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
                  .cell(CellIndex.indexByColumnRow(
                      columnIndex: i, rowIndex: rowIndex))
                  .value =
              row[i] != null
                  ? TextCellValue(row[i].toString())
                  : TextCellValue('');
        }
        rowIndex++;
      }

      // Encode Excel file
      List<int>? fileBytes = excel.encode();
      if (fileBytes == null) {
        throw Exception("Failed to encode Excel file.");
      }

      // Create a Blob from the bytes
      final blob = html.Blob([fileBytes],
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

      // Generate a download URL
      final url = html.Url.createObjectUrlFromBlob(blob);

      // Create an anchor element and trigger the download
      final anchor = html.document.createElement('a') as html.AnchorElement
        ..href = url
        ..style.display = 'none'
        ..download = fileName;
      html.document.body!.children.add(anchor);

      // Trigger the download
      anchor.click();

      // Clean up by removing the anchor and revoking the URL
      anchor.remove();
      html.Url.revokeObjectUrl(url);
    } catch (e) {
      throw Exception('Error exporting to Excel: $e');
    }
  }
}
