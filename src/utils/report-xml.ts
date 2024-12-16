import ExcelJS from "exceljs";
import { format, eachDayOfInterval, isBefore, parseISO } from "date-fns";
import { es } from "date-fns/locale";
import { AsistenciaData } from "../domain/interface/employee-attendance.interface";
import { getMonthName } from "./getMonthName";
import fs from 'fs';
import path from 'path';

// Obtener la ruta al archivo settings.json
const settingsFilePath = path.join(__dirname, "..", "../src/data/settings.json");
// Leer y parsear el archivo settings.json
const settingsData = fs.readFileSync(settingsFilePath, 'utf8');
const settings = JSON.parse(settingsData);



function getExcelColumnLetter(columnNumber: number): string {
  let columnLetter = '';
  while (columnNumber > 0) {
    let remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

const createHeader = (
  sheet: ExcelJS.Worksheet,
  totalColumns: number,
  monthName: string
) => {
  // Obtener la letra de la última columna
  const lastColumnLetter = getExcelColumnLetter(totalColumns + 5);

  // Encabezado principal
  sheet.mergeCells(`A1:${lastColumnLetter}1`);
  sheet.getCell("A1").value = "TALLER DE ESTRUCTURAS METÁLICAS";
  sheet.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getCell("A1").font = { size: 16, bold: true };

  // Subtítulo
  sheet.mergeCells(`A2:${lastColumnLetter}2`);
  sheet.getCell("A2").value = "CONTROL DE HORAS Y ASISTENCIA LABORAL";
  sheet.getCell("A2").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getCell("A2").font = { size: 14, bold: true };

  // Nombre del mes
  sheet.mergeCells(`A3:${lastColumnLetter}3`);
  sheet.getCell("A3").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getCell("A3").font = { size: 12, bold: true };
};


const createColumnHeaders = (
  sheet: ExcelJS.Worksheet,
  start: Date,
  daysArray: string[],
  weekdayArray: string[],
  // extraColumns: any[]
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

  let colIndex = endDayCol + 1;

  // Configurar anchos de columnas
  sheet.getColumn(1).width = 6; // 'N'
  sheet.getColumn(2).width = 10; // 'CÓDIGO'
  sheet.getColumn(3).width = 30; // 'NOMBRES'
  sheet.getColumn(4).width = 15; // 'DEPARTAMENTO'
  sheet.getColumn(5).width = 8; // 'DÍAS'

  for (let i = startDayCol; i <= endDayCol; i++) {
    sheet.getColumn(i).width = 3; 
  }

  colIndex = endDayCol + 1;
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
  const extraColumnsCount = 12; // Ajusta este valor según tus necesidades

  // 1. Convertir el objeto en un arreglo de [idEmpleado, Empleado[]]
  const entries = Object.entries(data);

  // 2. Ordenar las entradas por Departamento, Emp_lastname y Emp_firstname
  entries.sort(([, aRecords], [, bRecords]) => {
    const a = aRecords[0];
    const b = bRecords[0];

    const deptCompare = a.Departamento.localeCompare(b.Departamento);
    if (deptCompare !== 0) return deptCompare;

    const lastNameCompare = a.Emp_lastname.localeCompare(b.Emp_lastname);
    if (lastNameCompare !== 0) return lastNameCompare;

    return a.Emp_firstname.localeCompare(b.Emp_firstname);
  });

  // 3. Iterar sobre las entradas ya ordenadas
  entries.forEach(([ , employeeRecords ], index) => {
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
          if (Number(record.PaycodeID) === 11) {
            // Enfermedad
            horasPorDia[dayIndex] = settings.sickLeaveSymbol;
          } else if (Number(record.PaycodeID) === 12) {
            // Vacaciones
            horasPorDia[dayIndex] = "V";
          } else if (record.HI === "HI") {
            // Solo entrada
            horasPorDia[dayIndex] = settings.entranceOnlySymbol;
          } else if (record.HS === "HS") {
            // Solo salida
            horasPorDia[dayIndex] = settings.exitOnlySymbol;
          } else if (record.TipoZ === "Z") {
            // Ausencia
            horasPorDia[dayIndex] = settings.absenceSymbol;
          } else {
            // Total de horas
            horasPorDia[dayIndex] =
            record.TotalHorasRedondeadas !== null && record.TotalHorasRedondeadas !== 0
              ? record.TotalHorasRedondeadas.toString()
              : settings.entranceOnlySymbol;
          }
        }
      }
    });

    const rowValues = [
      index + 1, // 'N'
      empleado.Numero || "",
      empleado.Nombre,
      empleado.Departamento || "",
      daysArray.length,
      ...horasPorDia,
      ...Array(extraColumnsCount).fill(""),
    ];
    const row = sheet.addRow(rowValues);

    // Aplicar estilos a la fila
    row.eachCell((cell, colNumber) => {
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
        bottom: { style: "thin" },
      };

      // Colorear celdas según el valor
      if (colNumber > 5 && colNumber <= 5 + daysArray.length) {
        if (cell.value === "HI") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF" }, // Ajusta el color
          };
        } else if (cell.value === "HS") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF" }, // Ajusta el color
          };
        } else if (cell.value === "Z") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF" }, // Ajusta el color
          };
        } else if (cell.value === "E") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFC7CE" }, // Ajusta el color
          };
        } else if (cell.value === "V") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF00" }, // Ajusta el color
          };
        }
      }
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

  createHeader(sheet, eachDayOfInterval({ start, end }).length, getMonthName(start.getMonth() + 1));

  const dateRange = eachDayOfInterval({ start, end });
  const daysArray = dateRange.map((date) => format(date, "d", { locale: es }));
  const weekdayArray = dateRange.map((date) =>
    format(date, "EEEEE", { locale: es }).toUpperCase()
  );


  createColumnHeaders(sheet, start, daysArray, weekdayArray);

  styleHeaders(sheet);
  
  processEmployeeData(sheet, data, dateRange, daysArray);

  return workbook.xlsx.writeBuffer();
};
