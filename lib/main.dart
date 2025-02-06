import 'dart:io';
import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel Viewer Demo',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: const ExcelViewer(filePath: 'assets/existing_excel_file.xlsx'),
    );
  }
}

class ExcelViewer extends StatefulWidget {
  final String filePath;

  const ExcelViewer({super.key, required this.filePath});

  @override
  State<ExcelViewer> createState() => _ExcelViewerState();
}

class _ExcelViewerState extends State<ExcelViewer> {
  List<List<dynamic>> excelData = [];

  @override
  void initState() {
    super.initState();
    readExcel();
  }

  void readExcel() async {
    try {
      ByteData data = await rootBundle.load(widget.filePath);
      var bytes =
          data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
      var excel = Excel.decodeBytes(bytes);

      List<List<dynamic>> rows = [];

      for (var table in excel.tables.keys) {
        for (var row in excel.tables[table]!.rows) {
          rows.add(row.map((cell) => cell?.value ?? '').toList());
        }
      }
      setState(() {
        excelData = rows;
      });
    } catch (e) {
      debugPrint("reading error Excel: $e");
      setState(() {
        excelData = [];
      });
    }
  }

  @override
  Widget build(BuildContext context) {
    if (excelData.isEmpty) {
      return Scaffold(
        appBar: AppBar(title: const Text("Excel Viewer")),
        body: const Center(child: CircularProgressIndicator()),
      );
    }

    final headers = excelData.first;
    final dataRows = excelData.sublist(1);

    return Scaffold(
      appBar: AppBar(title: const Text("Excel Viewer")),
      body: SingleChildScrollView(
        scrollDirection: Axis.vertical,
        child: SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          child: DataTable(
            columns: headers
                .map(
                  (header) => DataColumn(
                    label: Padding(
                      padding: const EdgeInsets.all(8.0),
                      child: Text(
                        header.toString(),
                        style: const TextStyle(fontStyle: FontStyle.italic),
                      ),
                    ),
                  ),
                )
                .toList(),
            rows: dataRows
                .map(
                  (row) => DataRow(
                    cells: row
                        .map(
                          (cell) => DataCell(
                            Padding(
                              padding: const EdgeInsets.all(8.0),
                              child: Text(cell.toString()),
                            ),
                          ),
                        )
                        .toList(),
                  ),
                )
                .toList(),
          ),
        ),
      ),
    );
  }
}
