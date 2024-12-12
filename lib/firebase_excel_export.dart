import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:excel/excel.dart';

class FirebaseExcelExporter {
  /// Exports a Firestore collection to an Excel file.
  ///
  /// [collectionName] - The name of the Firestore collection to fetch data from.
  /// [headers] - The list of headers for the Excel file.
  /// [mapper] - A function that maps Firestore document data to a list of values for each row.
  /// [fileName] - The name of the resulting Excel file (default: 'export.xlsx').
  static Future<Excel> exportToExcel({
    required String collectionName,
    required List<String> headers,
    required List<dynamic> Function(Map<String, dynamic>) mapper,
  }) async {
    try {
      // Fetch data from Firestore
      QuerySnapshot snapshot =
      await FirebaseFirestore.instance.collection(collectionName).get();

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

      // Return the Excel object for further use
      return excel;
    } catch (e) {
      throw Exception('Error exporting to Excel: $e');
    }
  }
}
