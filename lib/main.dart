import 'package:flutter/material.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;
import 'dart:convert';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: MyHomePage(),
    );
  }
}

class DataSale {
  final int order_id;
  final String dc_number;
  final String V_number;
  final String item;
  final String qty;
  final String rate;
  final String total_amount;
  final String order_date;

  DataSale({required this.order_id,required this.dc_number,required this.V_number,required this.item,required this.qty,required this.rate,required this.total_amount,required this.order_date});

}


class MyHomePage extends StatefulWidget {
  @override
  _MyHomePageState createState() => _MyHomePageState();
}

List<ExcelDataRow> _buildCustomersDataRows() {
  List<ExcelDataRow> excelDataRows = <ExcelDataRow>[];
  final List<DataSale> reports = _getRecords();

  excelDataRows = reports.map<ExcelDataRow>((DataSale dataRow) {
    return ExcelDataRow(cells: <ExcelDataCell>[
      ExcelDataCell(columnHeader: 'Order Number', value: dataRow.order_id),
      ExcelDataCell(columnHeader: 'DC Number', value: dataRow.dc_number),
      ExcelDataCell(columnHeader: 'Vehicle Number', value: dataRow.V_number),
      ExcelDataCell(columnHeader: 'Item', value: dataRow.item),
      ExcelDataCell(columnHeader: 'Quantity', value: dataRow.qty),
      ExcelDataCell(columnHeader: 'Rate', value: dataRow.rate),
      ExcelDataCell(columnHeader: 'Total Amount', value: dataRow.total_amount),
      ExcelDataCell(columnHeader: 'Order Date', value: dataRow.order_date)
    ]);
  }).toList();

  return excelDataRows;
}


List<DataSale> _getRecords() {

  final List<DataSale> reports = <DataSale>[];
  reports.add(DataSale(order_id: 1, dc_number: "12346456", V_number: "RIK2764", item: "BestWay", qty: "100", rate: "3400", total_amount: "144000", order_date: "12-12-2021"));
  reports.add(DataSale(order_id: 2, dc_number: "67878555", V_number: "RIK6734", item: "BestWay", qty: "60", rate: "83240", total_amount: "44000", order_date: "12-12-2021"));
  reports.add(DataSale(order_id: 3, dc_number: "76867856", V_number: "RI6734", item: "BestWay", qty: "150", rate: "3430", total_amount: "3300", order_date: "12-12-2021"));
  reports.add(DataSale(order_id: 4, dc_number: "56756745", V_number: "RIK6734", item: "BestWay", qty: "540", rate: "82340", total_amount: "3300", order_date: "12-12-2021"));
  reports.add(DataSale(order_id: 5, dc_number: "12567345", V_number: "RIK74", item: "BestWay", qty: "5400", rate: "803240", total_amount: "133000", order_date: "12-12-2021"));

  return reports;
}


class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child:
        ElevatedButton(child: Text('Create Excel'), onPressed: createExcel),
      ),
    );
  }

  Future<void> createExcel() async {



    final Workbook workbook = Workbook();

    final Style headingStyle = workbook.styles.add('HeadingStyle');
    headingStyle.bold = true;
    headingStyle.hAlign = HAlignType.center;
    headingStyle.wrapText = true;

    final Style bodyStyle = workbook.styles.add('BodyStyle');
    bodyStyle.bold = false;
    bodyStyle.hAlign = HAlignType.center;
    bodyStyle.wrapText = true;


    final Worksheet sheet = workbook.worksheets[0];

    final List<ExcelDataRow> dataRows = _buildCustomersDataRows();

// Import the Data Rows in to Worksheet.
    sheet.importData(dataRows, 1, 1);

    sheet.getRangeByName('A1:H1').cellStyle = headingStyle;
    sheet.getRangeByName('A1:H1').columnWidth = 21;

    sheet.getRangeByName('A2:H100').cellStyle = bodyStyle;



    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    File('ApplyGlobalStyle.xlsx').writeAsBytes(bytes);

    if (kIsWeb) {
      AnchorElement(
          href:
          'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Output.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName =
      Platform.isWindows ? '$path\\Output.xlsx' : '$path/Output.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
