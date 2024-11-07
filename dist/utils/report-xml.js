"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createExcelReport = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const getMonthName_1 = require("./getMonthName");
const createExcelReport = (mes, year, data) => __awaiter(void 0, void 0, void 0, function* () {
    const workbook = new exceljs_1.default.Workbook();
    const sheet = workbook.addWorksheet("Asistencia");
    // Encabezado general
    sheet.mergeCells("A1:AM1");
    sheet.getCell("A1").value = "TALLER DE ESTRUCTURAS METÁLICAS";
    sheet.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };
    sheet.getCell("A1").font = { size: 16, bold: true };
    sheet.mergeCells("A2:AM2");
    sheet.getCell("A2").value = "CONTROL DE HORAS Y ASISTENCIA LABORAL";
    sheet.getCell("A2").alignment = { vertical: "middle", horizontal: "center" };
    sheet.getCell("A2").font = { size: 14, bold: true };
    // Genera días dinámicamente basado en el mes y año
    const daysInMonth = new Date(year, mes, 0).getDate();
    const daysArray = Array.from({ length: daysInMonth }, (_, i) => i + 1);
    const weekdayArray = daysArray.map((day) => {
        const date = new Date(year, mes - 1, day);
        const weekdays = ["D", "L", "M", "M", "J", "V", "S"];
        return weekdays[date.getDay()];
    });
    // Encabezado de columnas
    sheet.getRow(4).values = [
        "COD",
        "NOMBRES",
        "DEPARTAMENTO",
        "DÍAS",
        (0, getMonthName_1.getMonthName)(mes),
    ];
    sheet.getRow(5).values = ["", "", "", "", ...daysArray];
    sheet.getRow(6).values = ["", "", "", "", ...weekdayArray];
    // Fusiona las celdas para encabezados
    sheet.mergeCells("A4:A6"); // COD
    sheet.mergeCells("B4:B6"); // NOMBRES
    sheet.mergeCells("C4:C6"); // DEPARTAMENTO
    sheet.mergeCells("D4:D6"); // DÍAS
    // Usa switch para determinar la última columna de la fusión según los días del mes
    switch (daysInMonth) {
        case 28:
            sheet.mergeCells("E4:AF4"); // 28 días
            break;
        case 29:
            sheet.mergeCells("E4:AG4"); // 29 días
            break;
        case 30:
            sheet.mergeCells("E4:AH4"); // 30 días
            break;
        case 31:
            sheet.mergeCells("E4:AI4"); // 31 días
            break;
        default:
            throw new Error("Número de días en el mes no válido");
    }
    // Ajusta el ancho de las columnas
    sheet.getColumn(1).width = 6; // COD
    sheet.getColumn(2).width = 30; // NOMBRES
    sheet.getColumn(3).width = 15; // DEPARTAMENTO
    sheet.getColumn(4).width = 8; // DÍAS
    for (let i = 5; i < 5 + daysInMonth; i++) {
        sheet.getColumn(i).width = 3; // Columnas de días
    }
    // Estilo de encabezados
    [4, 5, 6].forEach((rowNum) => {
        sheet.getRow(rowNum).eachCell((cell) => {
            cell.font = { bold: true };
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                right: { style: "thin" },
                bottom: { style: "thin" },
            };
            if (cell.value === "S" || cell.value === "D") {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFF00" },
                };
            }
        });
    });
    // Procesar los datos de cada empleado en `data`
    Object.values(data).forEach((employeeRecords) => {
        if (employeeRecords.length === 0)
            return;
        // Obtener los detalles del empleado y preparar una fila con las horas trabajadas por día
        const empleado = employeeRecords[0];
        const horasPorDia = Array(daysInMonth).fill(""); // Inicializa el array para días del mes
        employeeRecords.forEach((record) => {
            const day = new Date(record.Fecha).getDate() - 1;
            horasPorDia[day] = record.TotalHorasRedondeadas || ""; // Coloca las horas trabajadas o vacío si no hay
        });
        const rowValues = [
            empleado.Numero || "",
            empleado.Nombre,
            empleado.Departamento,
            daysInMonth,
            ...horasPorDia,
        ];
        const row = sheet.addRow(rowValues);
        // Estilo para las celdas de datos
        row.eachCell((cell) => {
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                right: { style: "thin" },
                bottom: { style: "thin" },
            };
        });
    });
    // Guardar el archivo
    return workbook.xlsx.writeBuffer();
});
exports.createExcelReport = createExcelReport;
