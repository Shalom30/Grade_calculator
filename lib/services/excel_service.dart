// ============================================================
//  FILE: lib/services/excel_service.dart
// ============================================================

import 'dart:io';
import 'package:excel/excel.dart';
import '../models/student.dart';

class ExcelService {
  // ── READ ────────────────────────────────────────────────────
  static Future<List<Student>> readStudents(String filePath) async {
    final file = File(filePath);
    if (!await file.exists()) {
      throw Exception('File not found: $filePath');
    }

    final bytes = await file.readAsBytes();
    final excel = Excel.decodeBytes(bytes);

    final sheet = excel.tables[excel.tables.keys.first];
    if (sheet == null) {
      throw Exception('The Excel file has no sheets.');
    }

    final List<Student> students = [];

    for (int i = 1; i < sheet.maxRows; i++) {
      final row = sheet.row(i);

      final name     = _cellString(row, 0);
      final caMark   = _cellDouble(row, 1);
      final examMark = _cellDouble(row, 2);

      if (name.isEmpty) continue;

      if (caMark < 0 || caMark > 40) {
        throw Exception('CA mark for "$name" is $caMark — must be 0–40.');
      }
      if (examMark < 0 || examMark > 100) {
        throw Exception('Exam mark for "$name" is $examMark — must be 0–100.');
      }

      students.add(Student(name: name, caMark: caMark, examMark: examMark));
    }

    if (students.isEmpty) {
      throw Exception(
        'No student data found.\n'
        'Column A: Student Name\n'
        'Column B: CA Mark (out of 40)\n'
        'Column C: Exam Mark (out of 100)',
      );
    }

    return students;
  }

  // ── WRITE ───────────────────────────────────────────────────
  static Future<void> writeResults(
    List<Student> students,
    String outputPath,
  ) async {
    final excel = Excel.createExcel();
    final sheet = excel['Results'];

    final headers = [
      'Student Name', 'CA Mark', 'Exam Mark',
      'Final Mark', 'Grade', 'GPA',
    ];

    for (int col = 0; col < headers.length; col++) {
      final cell = sheet.cell(
        CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0),
      );
      cell.value = TextCellValue(headers[col]);
      cell.cellStyle = CellStyle(
        bold: true,
        backgroundColorHex: ExcelColor.fromHexString('#1A1A2E'),
        fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
      );
    }

    for (int i = 0; i < students.length; i++) {
      final s = students[i];
      final rowIndex = i + 1;

      void setCell(int col, CellValue value) {
        sheet
            .cell(CellIndex.indexByColumnRow(
              columnIndex: col,
              rowIndex: rowIndex,
            ))
            .value = value;
      }

      setCell(0, TextCellValue(s.name));
      setCell(1, DoubleCellValue(s.caMark));
      setCell(2, DoubleCellValue(s.examMark));
      setCell(3, DoubleCellValue(
        double.parse(s.finalMark.toStringAsFixed(2)),
      ));
      setCell(4, TextCellValue(s.grade));
      setCell(5, DoubleCellValue(s.gpa));
    }

    for (int col = 0; col < headers.length; col++) {
      sheet.setColumnWidth(col, 18);
    }

    final outputBytes = excel.save();
    if (outputBytes == null) {
      throw Exception('Failed to encode Excel file.');
    }

    await File(outputPath).writeAsBytes(outputBytes);
  }

  // ── Private helpers ─────────────────────────────────────────
  static String _cellString(List<Data?> row, int col) {
    if (col >= row.length) return '';
    return row[col]?.value?.toString().trim() ?? '';
  }

  static double _cellDouble(List<Data?> row, int col) {
    if (col >= row.length) return 0.0;
    final val = row[col]?.value;
    if (val == null) return 0.0;
    if (val is IntCellValue)    return val.value.toDouble();
    if (val is DoubleCellValue) return val.value;
  if (val is TextCellValue)   return double.tryParse(val.value.toString()) ?? 0.0;
    return double.tryParse(val.toString()) ?? 0.0;
  }
}