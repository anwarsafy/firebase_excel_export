# Firebase Excel Export

A Flutter package to export Firestore collections to Excel with a simple and customizable approach. Perfect for Flutter web applications!

## Features
- Export Firestore data directly to `.xlsx` files.
- Fully customizable headers and data mapping.
- Easy integration with Firestore.
- Designed for Flutter web applications with seamless file download functionality.

## Installation

Add the package to your `pubspec.yaml` file:

```yaml
dependencies:
  firebase_excel_export: ^0.0.1
```

Run the following command to fetch the package:

```bash
flutter pub get
```

## Usage

### Basic Example

```dart
import 'package:firebase_excel_export/firebase_excel_export.dart';

Future<void> exportUsers() async {
  await FirebaseExcelExporter.exportToExcel(
    collectionName: 'users',
    headers: ['Email', 'Name', 'Phone Number'], // Customize your headers
    mapper: (data) => [
      data['email'],       // Map Firestore fields to the headers
      data['name'],
      data['phoneNumber'],
    ],
    fileName: 'users_export.xlsx', // Optional: Customize file name
  );
}
```

### Parameters
- `collectionName` *(required)*: The Firestore collection to export.
- `headers` *(required)*: A list of custom headers for the Excel file.
- `mapper` *(required)*: A function to map Firestore document fields to header columns.
- `fileName` *(optional)*: The desired file name for the Excel file (default: `export.xlsx`).

### Advanced Example with Query Filtering

```dart
Future<void> exportFilteredUsers() async {
  await FirebaseExcelExporter.exportToExcel(
    collectionName: 'users',
    headers: ['Email', 'Name', 'Status'], // Define custom headers
    mapper: (data) => [
      data['email'],       // Map data fields dynamically
      data['name'],
      data['status'],
    ],
    queryBuilder: (query) => query.where('status', isEqualTo: 'active'), // Apply Firestore query
    fileName: 'active_users_export.xlsx', // Define a custom file name
  );
}
```

## Key Features
1. **Dynamic Headers**: Define custom headers tailored to your Firestore data.
2. **Data Mapping**: Map fields dynamically to headers for total flexibility.
3. **Web-Optimized**: Built for Flutter web to enable seamless downloads.

## Compatibility
This package is designed for **Flutter web** and depends on:
- `cloud_firestore`
- `firebase_core`
- `excel`
- `universal_html`

Ensure these dependencies are added to your `pubspec.yaml`.

## Support
If you encounter any issues or have feature requests, feel free to [open an issue](https://github.com/your-repo/firebase_excel_export/issues) on the GitHub repository.

## License
This package is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.