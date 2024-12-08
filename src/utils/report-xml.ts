import ExcelJS from "exceljs";
import { format, eachDayOfInterval, isBefore, parseISO } from "date-fns";
import { es } from "date-fns/locale";
import { AsistenciaData } from "../domain/interface/employee-attendance.interface";
import { getMonthName } from "./getMonthName";

const createHeader = (sheet: ExcelJS.Worksheet) => {
  sheet.mergeCells("A1:AM1");
  sheet.getCell("A1").value = "TALLER DE ESTRUCTURAS METÁLICAS";
  sheet.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getCell("A1").font = { size: 16, bold: true };

  sheet.mergeCells("A2:AM2");
  sheet.getCell("A2").value = "CONTROL DE HORAS Y ASISTENCIA LABORAL";
  sheet.getCell("A2").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getCell("A2").font = { size: 14, bold: true };
};

const createColumnHeaders = (
  sheet: ExcelJS.Worksheet,
  start: Date,
  daysArray: string[],
  weekdayArray: string[],
  extraColumns: any[]
) => {
  const startDayCol = 6; // Columna 'F'
  const endDayCol = startDayCol + daysArray.length - 1;

  // Insertar filas si es necesario
  // Asegurarse de que haya al menos 5 filas antes de agregar encabezados
  if (sheet.rowCount < 5) {
    sheet.addRows(Array(5 - sheet.rowCount).fill([]));
  }

  // Índices de filas de encabezado
  const headerRow1Index = 5; // Fila para los encabezados principales
  const headerRow2Index = 6; // Fila para los días del mes
  const headerRow3Index = 7; // Fila para los días de la semana

  // Fusionar celdas para los encabezados principales en columnas 1 a 5
  const headers = ["N", "CÓDIGO", "NOMBRES", "DEPARTAMENTO", "DÍAS"];
  for (let col = 1; col <= 5; col++) {
    sheet.mergeCells(headerRow1Index, col, headerRow3Index, col);
    const cell = sheet.getCell(headerRow1Index, col);
    cell.value = headers[col - 1];
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.font = { bold: true };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
      bottom: { style: "thin" },
    };
  }

  // Fila 1: Nombre del mes (fusionado encima de los días)
  sheet.mergeCells(headerRow1Index, startDayCol, headerRow1Index, endDayCol);
  const monthCell = sheet.getCell(headerRow1Index, startDayCol);
  monthCell.value = getMonthName(start.getMonth() + 1).toUpperCase();
  monthCell.font = { bold: true };
  monthCell.alignment = { horizontal: "center", vertical: "middle" };

  // Fila 2: Días del mes
  for (let i = 0; i < daysArray.length; i++) {
    const col = startDayCol + i;
    sheet.mergeCells(headerRow2Index, col, headerRow2Index, col);
    const cell = sheet.getCell(headerRow2Index, col);
    cell.value = daysArray[i];
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.font = { bold: true };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
      bottom: { style: "thin" },
    };
  }

  // Fila 3: Días de la semana
  for (let i = 0; i < weekdayArray.length; i++) {
    const col = startDayCol + i;
    const cell = sheet.getCell(headerRow3Index, col);
    cell.value = weekdayArray[i];
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.font = { bold: true };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
      bottom: { style: "thin" },
    };
  }

  // Manejo de columnas extra
  let colIndex = endDayCol + 1;
  extraColumns.forEach((col) => {
    const colSpan = col.span || 1;
    const startCol = colIndex;
    const endCol = colIndex + colSpan - 1;

    if (colSpan > 1) {
      // Fusionar horizontalmente en headerRow2Index
      sheet.mergeCells(headerRow2Index, startCol, headerRow2Index, endCol);
      sheet.getCell(headerRow2Index, startCol).value = col.header;
      sheet.getCell(headerRow2Index, startCol).alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      sheet.getCell(headerRow2Index, startCol).font = { bold: true };
      // Subencabezados en la fila 3
      for (let i = 0; i < colSpan; i++) {
        const subHeaderCell = sheet.getCell(headerRow3Index, startCol + i);
        subHeaderCell.value = col.subHeaders && col.subHeaders[i] ? col.subHeaders[i] : "";
        subHeaderCell.alignment = { horizontal: "center", vertical: "middle" };
        subHeaderCell.font = { bold: true };
        subHeaderCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
          bottom: { style: "thin" },
        };
      }
    } else {
      // Fusionar verticalmente desde headerRow2Index a headerRow3Index
      sheet.mergeCells(headerRow2Index, startCol, headerRow3Index, startCol);
      const cell = sheet.getCell(headerRow2Index, startCol);
      cell.value = col.header;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { bold: true };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };
    }

    colIndex = endCol + 1;
  });

  // Configurar anchos de columnas
  sheet.getColumn(1).width = 6; // 'N'
  sheet.getColumn(2).width = 10; // 'CÓDIGO'
  sheet.getColumn(3).width = 30; // 'NOMBRES'
  sheet.getColumn(4).width = 15; // 'DEPARTAMENTO'
  sheet.getColumn(5).width = 8; // 'DÍAS'

  for (let i = startDayCol; i <= endDayCol; i++) {
    sheet.getColumn(i).width = 3; // Columnas de días
  }

  // Configurar anchos de columnas adicionales
  colIndex = endDayCol + 1;
  extraColumns.forEach((col) => {
    for (let i = 0; i < col.span; i++) {
      const currentCol = sheet.getColumn(colIndex);
      if (
        ["FALTAS", "SALIDAS", "VACACIÓN", "ENFERMEDAD"].includes(col.header)
      ) {
        currentCol.width = 8; // Ancho para columnas verticales
      } else {
        currentCol.width = 15; // Ancho estándar
      }
      colIndex++;
    }
  });
};

const styleHeaders = (sheet: ExcelJS.Worksheet) => {
  const verticalHeaders = ["FALTAS", "SALIDAS", "VACACIÓN", "ENFERMEDAD"];
  const verticalHeaderColumns: number[] = [];

  // Ajustar la altura de las filas de encabezado
  sheet.getRow(5).height = 20; // Fila para el nombre del mes
  sheet.getRow(6).height = 20; // Fila de días del mes
  sheet.getRow(7).height = 100; // Fila de días de la semana y vertical headers

  [5, 6, 7].forEach((rowNum) => {
    sheet.getRow(rowNum).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const cellValue = cell.value as string;
      const isVerticalHeader = verticalHeaders.includes(cellValue);

      cell.font = { bold: true };
      cell.alignment = {
        horizontal: "center",
        vertical: "middle",
        textRotation: isVerticalHeader ? 90 : 0,
        wrapText: false,
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };

      if (isVerticalHeader) {
        verticalHeaderColumns.push(colNumber);
      }

      // Resaltar sábados y domingos en la fila de días de la semana
      if (rowNum === 7 && (cellValue === "S" || cellValue === "D")) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF00" },
        };
      }
    });
  });

  // Ajustar el ancho de las columnas verticales
  verticalHeaderColumns.forEach((colNumber) => {
    sheet.getColumn(colNumber).width = 8; // Ajusta este valor según necesites
  });
};


const processEmployeeData = (
  sheet: ExcelJS.Worksheet,
  data: AsistenciaData,
  dateRange: Date[],
  daysArray: string[]
) => {
  const extraColumnsCount = 12; // Actualiza este valor según el número de columnas extra
  const startingRowIndex = 8; // La fila donde empiezan los datos de empleados

  Object.values(data).forEach((employeeRecords, index) => {
    if (employeeRecords.length === 0) return;

    const empleado = employeeRecords[0];
    const horasPorDia = Array(daysArray.length).fill("");

    employeeRecords.forEach((record) => {
      const recordDate = parseISO(record.Fecha);
      if (
        recordDate >= dateRange[0] &&
        recordDate <= dateRange[dateRange.length - 1]
      ) {
        const dayIndex = dateRange.findIndex(
          (date) =>
            format(date, "yyyy-MM-dd") === format(recordDate, "yyyy-MM-dd")
        );
        if (dayIndex >= 0) {
          horasPorDia[dayIndex] = record.TotalHorasRedondeadas || "";
        }
      }
    });

    const rowValues = [
      index + 1, // 'N'
      empleado.Numero || "",
      empleado.Nombre,
      empleado.Departamento,
      daysArray.length,
      ...horasPorDia,
      // Placeholders para columnas extra
      ...Array(extraColumnsCount).fill(""),
    ];
    const row = sheet.addRow(rowValues);

    row.eachCell((cell) => {
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };
    });
  });
};

export const createExcelReport = async (
  startDate: string,
  endDate: string,
  data: AsistenciaData
) => {
  const start = parseISO(startDate);
  const end = parseISO(endDate);

  if (isBefore(end, start)) {
    throw new Error(
      "La fecha de fin no puede ser anterior a la fecha de inicio."
    );
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Asistencia");

  createHeader(sheet);

  const dateRange = eachDayOfInterval({ start, end });
  const daysArray = dateRange.map((date) => format(date, "d", { locale: es }));
  const weekdayArray = dateRange.map((date) =>
    format(date, "EEEEE", { locale: es }).toUpperCase()
  );

  // Obtener el nombre del mes
  const monthName = getMonthName(start.getMonth() + 1);

  const extraColumns = [
    { header: "# HORAS EXTRAS", subHeaders: ["50%", "100%"], span: 2 },
    { header: "BÁSICA", span: 1 },
    { header: "SUELDO BÁSICO", span: 1 },
    { header: "FALTAS", span: 1 },
    { header: "SALIDAS", span: 1 },
    { header: "VACACIÓN", span: 1 },
    { header: "ENFERMEDAD", span: 1 },
    {
      header: "MONTO POR HORAS EXTRAS",
      subHeaders: ["50%", "100%"],
      span: 2,
    },
    { header: "SUBTOTAL", span: 1 },
  ];

  createColumnHeaders(
    sheet,
    start,
    daysArray,
    weekdayArray,
    extraColumns,
  );

  styleHeaders(sheet);
  processEmployeeData(sheet, data, dateRange, daysArray);

  return workbook.xlsx.writeBuffer();
};
