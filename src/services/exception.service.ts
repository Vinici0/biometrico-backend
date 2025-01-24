// src/services/exception.service.ts

import PdfPrinter from "pdfmake";
import path from "path";
import Sequelize from "sequelize";
import sequelize from "../config/database";
import { TDocumentDefinitions } from "pdfmake/interfaces";
import { resultsInit } from "../data/init";
import ExcelJS from "exceljs";
import { format, parseISO } from 'date-fns';
import { ExceptionRecord, Attendance } from "../domain/interface/exception.interface";
import fs from "fs";

export class ExceptionService {
  private printer: PdfPrinter;

  constructor() {
    // Definir la ruta relativa a la carpeta de fuentes
    const fontsPath = path.join(__dirname, "..", "fonts/fonts");

    // Inicializar PdfPrinter con las rutas de las fuentes
    this.printer = new PdfPrinter({
      Roboto: {
        normal: path.join(fontsPath, "Roboto-Regular.ttf"),
        bold: path.join(fontsPath, "Roboto-Regular.ttf"),
        italics: path.join(fontsPath, "Roboto-Italic.ttf"),
        bolditalics: path.join(fontsPath, "Roboto-MediumItalic.ttf"),
      },
    });
  }
  public async getExceptionReport(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30"
  ): Promise<any> {
    try {
      // Datos de ejemplo
      const results = (await sequelize.query(
        ` 
      SELECT
          HE.emp_pin AS emp_pin,
          HE.emp_firstname AS emp_firstname,
          HE.emp_lastname AS emp_lastname,
          AEA.exception_date AS exception_date,
          AEA.starttime AS starttime,
          AEA.endtime AS endtime,
          ROUND((julianday(AEA.endtime) - julianday(AEA.starttime)) * 24, 2) AS total_horas_excepcion,
          CASE
              WHEN AP.pc_desc LIKE '%vacation%' THEN 'Vacaciones'
              WHEN AP.pc_desc LIKE '%sick%' THEN 'Enfermedad'
              ELSE AP.pc_desc
          END AS tipo_de_excepcion
      FROM
          att_paycode AP
      INNER JOIN
          att_exceptionassign AEA ON AP.id = AEA.paycode_id
      INNER JOIN
          hr_employee HE ON AEA.employee_id = HE.id
      WHERE
          AEA.exception_date BETWEEN :startDate AND :endDate
          AND HE.emp_active = 1
      ORDER BY
          HE.emp_pin;
        `,
        {
          replacements: { startDate, endDate },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as ExceptionRecord[];

      // Filtrar los resultados por rango de fechas si se proporcionan
      const filteredResults = this.filterResultsByDate(
        results,
        startDate,
        endDate
      );

      const documentDefinition = this.generateExceptionReport(filteredResults);

      return await this.createPdfBuffer(documentDefinition);
    } catch (error) {
      console.error("Error generando el reporte:", error);
      throw error;
    }
  }

  private filterResultsByDate(
    data: ExceptionRecord[],
    startDate: string,
    endDate: string
  ): ExceptionRecord[] {
    if (!startDate && !endDate) {
      return data;
    }

    const start = startDate ? new Date(startDate) : new Date("1970-01-01");
    const end = endDate ? new Date(endDate) : new Date("2100-12-31");

    return data.filter((record) => {
      const recordDate = new Date(record.exception_date);
      return recordDate >= start && recordDate <= end;
    });
  }

  private generateExceptionReport(
    data: ExceptionRecord[]
  ): TDocumentDefinitions {
    const formattedData = data.map((item) => [
      { text: item.emp_pin.toString(), alignment: "center" },
      { text: `${item.emp_firstname} ${item.emp_lastname}`, alignment: "left" },
      {
        text: new Date(item.exception_date).toLocaleDateString("es-ES"),
        alignment: "center",
      },
      {
        text: new Date(item.starttime).toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
          hour12: false,
        }),
        alignment: "center",
      },
      {
        text: new Date(item.endtime).toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
          hour12: false,
        }),
        alignment: "center",
      },
      { text: item.total_horas_excepcion.toString(), alignment: "center" },
      { text: item.tipo_de_excepcion, alignment: "center" },
    ]);

    const documentDefinition: TDocumentDefinitions = {
      pageSize: "A4",
      pageMargins: [40, 60, 40, 60],
      header: (currentPage: number, pageCount: number) => {
        if (currentPage === 1) {
          return {
            columns: [
              {
                image: path.join(__dirname, "../assets/assets/logo.png"),
                width: 70,
                margin: [40, 20, 0, 20],
              },
              {
                text: "Reporte de Excepciones",
                alignment: "center",
                fontSize: 18, // Tamaño de la fuente del encabezado
                bold: true,
                margin: [0, 25, 50, 0],
              },
            ],
          };
        } else {
          return null; // No mostrar encabezado en otras páginas
        }
      },
      footer: {
        columns: [
          {
            text: `Fecha de Reporte: ${new Date().toLocaleDateString("es-ES")}`,
            alignment: "left",
            margin: [40, 0, 0, 0],
          },
        ],
        margin: [0, 0, 0, 30],
      },
      content: [
        {
          table: {
            headerRows: 1,
            widths: ["10%", "20%", "15%", "13%", "13%", "14%", "15%"],
            body: [
              [
                { text: "ID", style: "tableHeader" },
                { text: "Nombre", style: "tableHeader" },
                { text: "Fecha", style: "tableHeader" },
                { text: "Hora Inicio", style: "tableHeader" },
                { text: "Hora Fin", style: "tableHeader" },
                { text: "Excepción", style: "tableHeader" },
                { text: "Tipo de Excepción", style: "tableHeader" },
              ],
              ...formattedData,
            ],
          },
          layout: {
            fillColor: (rowIndex: number) =>
              rowIndex === 0
                ? "#CCCCCC"
                : rowIndex % 2 === 0
                  ? "#F5F5F5"
                  : null, // Alternar colores
            paddingLeft: () => 6, // Reducir el espaciado interno de las celdas
            paddingRight: () => 6,
            paddingTop: () => 4,
            paddingBottom: () => 4,
          },
          margin: [0, 20, 0, 0], // Espacio alrededor de la tabla
        },
      ],
      styles: {
        tableHeader: {
          bold: true,
          fontSize: 10, // Reducir el tamaño de la fuente de los encabezados de la tabla
          color: "black",
        },
        subheader: {
          fontSize: 12, // Reducir el tamaño de la fuente de los subencabezados
          bold: true,
        },
      },
      defaultStyle: {
        font: "Roboto",
        fontSize: 8, // Reducir el tamaño de la fuente predeterminada
      },
    };

    return documentDefinition;
  }

  private createPdfBuffer(
    documentDefinition: TDocumentDefinitions
  ): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      try {
        const pdfDoc = this.printer.createPdfKitDocument(documentDefinition);
        const chunks: any[] = [];

        pdfDoc.on("data", (chunk) => {
          chunks.push(chunk);
        });

        pdfDoc.on("end", () => {
          const result = Buffer.concat(chunks);
          resolve(result);
        });

        pdfDoc.end();
      } catch (error) {
        reject(error);
      }
    });
  }

  public async getExceptionReportExcel(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30"
  ): Promise<ExcelJS.Workbook> {
    try {
      // Obtener los datos
      const results = (await sequelize.query(
        `
        SELECT
          HE.emp_pin AS emp_pin,
          HE.emp_firstname AS emp_firstname,
          HE.emp_lastname AS emp_lastname,
          AEA.exception_date AS exception_date,
          AEA.starttime AS starttime,
          AEA.endtime AS endtime,
          ROUND((julianday(AEA.endtime) - julianday(AEA.starttime)) * 24, 2) AS total_horas_excepcion,
          CASE
              WHEN AP.pc_desc LIKE '%vacation%' THEN 'Vacaciones'
              WHEN AP.pc_desc LIKE '%sick%' THEN 'Enfermedad'
              ELSE AP.pc_desc
          END AS tipo_de_excepcion
        FROM
          att_paycode AP
        INNER JOIN
          att_exceptionassign AEA ON AP.id = AEA.paycode_id
        INNER JOIN
          hr_employee HE ON AEA.employee_id = HE.id
        WHERE
          AEA.exception_date BETWEEN :startDate AND :endDate
          AND HE.emp_active = 1
        ORDER BY
          HE.emp_pin;
        `,
        {
          replacements: { startDate, endDate },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as ExceptionRecord[];

      // Filtrar los resultados por rango de fechas si se proporcionan
      const filteredResults = this.filterResultsByDate(
        results,
        startDate,
        endDate
      );

      // Crear el workbook y el worksheet
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Reporte de Excepciones");

      // Definir las columnas
      worksheet.columns = [
        { header: "ID", key: "emp_pin", width: 10 },
        { header: "Nombre", key: "nombre", width: 30 },
        { header: "Fecha", key: "exception_date", width: 15 },
        { header: "Hora Inicio", key: "starttime", width: 15 },
        { header: "Hora Fin", key: "endtime", width: 15 },
        {
          header: "Horas de Excepción",
          key: "total_horas_excepcion",
          width: 20,
        },
        { header: "Tipo de Excepción", key: "tipo_de_excepcion", width: 20 },
      ];

      // Agregar los datos
      filteredResults.forEach((record) => {
        worksheet.addRow({
          emp_pin: record.emp_pin,
          nombre: `${record.emp_firstname} ${record.emp_lastname}`,
          exception_date: new Date(record.exception_date).toLocaleDateString(
            "es-ES"
          ),
          starttime: new Date(record.starttime).toLocaleTimeString([], {
            hour: "2-digit",
            minute: "2-digit",
            hour12: false,
          }),
          endtime: new Date(record.endtime).toLocaleTimeString([], {
            hour: "2-digit",
            minute: "2-digit",
            hour12: false,
          }),
          total_horas_excepcion: record.total_horas_excepcion,
          tipo_de_excepcion: record.tipo_de_excepcion,
        });
      });

      // Estilos opcionales
      worksheet.getRow(1).font = { bold: true }; // Encabezados en negrita

      // Retornar el workbook
      return workbook;
    } catch (error) {
      console.error("Error generando el reporte de Excel:", error);
      throw error;
    }
  }

  public async generateAttendanceReportPDF(
    data: Attendance[]
  ): Promise<Buffer> {
    const formattedData = data.map((item) => [
      { text: item.id.toString(), alignment: "center" },
      { text: item.Numero, alignment: "center" },
      { text: item.Nombre, alignment: "left" },
      {
        text: new Date(item.Fecha).toLocaleDateString("es-ES"),
        alignment: "center",
      },
      { text: item.Entrada, alignment: "center" },
      { text: item.Salida || "", alignment: "center" },
      { text: item.Departamento, alignment: "center" },
      {
        text: item.TotalHorasRedondeadas?.toString() || "",
        alignment: "center",
      },
      { text: item.DiaSemana, alignment: "center" },
      { text: item.status, alignment: "center" },
    ]);

    const documentDefinition: TDocumentDefinitions = {
      pageSize: "A4",
      pageMargins: [40, 60, 40, 60],
      header: (currentPage: number, pageCount: number) => {
        if (currentPage === 1) {
          return {
            columns: [
              {
                image: path.join(__dirname, "..", "assets/assets", "logo.png"),
                width: 70,
                margin: [40, 20, 0, 20],
              },
              {
                text: "Reporte de Asistencia",
                alignment: "center",
                fontSize: 18,
                bold: true,
                margin: [0, 25, 50, 0],
              },
            ],
          };
        } else {
          return null;
        }
      },
      footer: {
        columns: [
          {
            text: `Fecha de Reporte: ${new Date().toLocaleDateString("es-ES")}`,
            alignment: "left",
            margin: [40, 0, 0, 0],
          },
        ],
        margin: [0, 0, 0, 30],
      },
      content: [
        {
          table: {
            headerRows: 1,
            widths: [
              "5%",
              "10%",
              "20%",
              "10%",
              "10%",
              "10%",
              "15%",
              "10%",
              "5%",
              "5%",
            ],
            body: [
              [
                { text: "ID", style: "tableHeader" },
                { text: "Número", style: "tableHeader" },
                { text: "Nombre", style: "tableHeader" },
                { text: "Fecha", style: "tableHeader" },
                { text: "Entrada", style: "tableHeader" },
                { text: "Salida", style: "tableHeader" },
                { text: "Departamento", style: "tableHeader" },
                { text: "Horas", style: "tableHeader" },
                { text: "Día", style: "tableHeader" },
                { text: "Estado", style: "tableHeader" },
              ],
              ...formattedData,
            ],
          },
          layout: {
            fillColor: (rowIndex: number) =>
              rowIndex === 0
                ? "#CCCCCC"
                : rowIndex % 2 === 0
                  ? "#F5F5F5"
                  : null,
            paddingLeft: () => 6,
            paddingRight: () => 6,
            paddingTop: () => 4,
            paddingBottom: () => 4,
          },
          margin: [0, 20, 0, 0],
        },
      ],
      styles: {
        tableHeader: {
          bold: true,
          fontSize: 10,
          color: "black",
        },
        subheader: {
          fontSize: 12,
          bold: true,
        },
      },
      defaultStyle: {
        font: "Roboto",
        fontSize: 8,
      },
    };

    return new Promise((resolve, reject) => {
      try {
        const pdfDoc = this.printer.createPdfKitDocument(documentDefinition);
        const chunks: any[] = [];

        pdfDoc.on("data", (chunk) => {
          chunks.push(chunk);
        });

        pdfDoc.on("end", () => {
          const result = Buffer.concat(chunks);
          resolve(result);
        });

        pdfDoc.end();
      } catch (error) {
        reject(error);
      }
    });
  }

  // public generateAttendanceReportPDFBuffer(
  //   data: Attendance[]
  // ): Promise<Buffer> {
  //   const formattedData = data.map((item) => [
  //     { text: item.Nombre, alignment: "left" },
  //     {
  //       text: new Date(item.Fecha).toLocaleDateString("es-ES"),
  //       alignment: "center",
  //     },
  //     { text: item.Entrada, alignment: "center" },
  //     { text: item.Salida || "", alignment: "center" },
  //     { text: item.Departamento, alignment: "center" },
  //     {
  //       text: item.TotalHorasRedondeadas?.toString() || "",
  //       alignment: "center",
  //     },
  //     { text: item.DiaSemana, alignment: "center" },
  //     { text: item.status, alignment: "center" },
  //   ]);

  //   const documentDefinition: TDocumentDefinitions = {
  //     pageSize: "A4",
  //     pageMargins: [40, 60, 40, 60],
  //     header: (currentPage: number, pageCount: number) => {
  //       if (currentPage === 1) {
  //         return {
  //           columns: [
  //             {
  //               image: path.join(__dirname, "..", "assets", "logo.png"),
  //               width: 70,
  //               margin: [40, 20, 0, 20],
  //             },
  //             {
  //               text: "Reporte de Asistencia",
  //               alignment: "center",
  //               fontSize: 18,
  //               bold: true,
  //               margin: [0, 25, 50, 0],
  //             },
  //           ],
  //         };
  //       } else {
  //         return null;
  //       }
  //     },
  //     footer: {
  //       columns: [
  //         {
  //           text: `Fecha de Reporte: ${new Date().toLocaleDateString("es-ES")}`,
  //           alignment: "left",
  //           margin: [40, 0, 0, 0],
  //         },
  //       ],
  //       margin: [0, 0, 0, 30],
  //     },
  //     content: [
  //       {
  //         table: {
  //           headerRows: 1,
  //           widths: ["25%", "10%", "10%", "15%", "10%", "10%", "10%", "10%"],
  //           body: [
  //             [
  //               { text: "Nombre", style: "tableHeader" },
  //               { text: "Fecha", style: "tableHeader" },
  //               { text: "Entrada", style: "tableHeader" },
  //               { text: "Salida", style: "tableHeader" },
  //               { text: "Departamento", style: "tableHeader" },
  //               { text: "Horas", style: "tableHeader" },
  //               { text: "Día", style: "tableHeader" },
  //               { text: "Estado", style: "tableHeader" },
  //             ],
  //             ...formattedData,
  //           ],
  //         },
  //         layout: {
  //           fillColor: (rowIndex: number) =>
  //             rowIndex === 0
  //               ? "#CCCCCC"
  //               : rowIndex % 2 === 0
  //               ? "#F5F5F5"
  //               : null,
  //           paddingLeft: () => 6,
  //           paddingRight: () => 6,
  //           paddingTop: () => 4,
  //           paddingBottom: () => 4,
  //         },
  //         margin: [0, 20, 0, 0],
  //       },
  //     ],
  //     styles: {
  //       tableHeader: {
  //         bold: true,
  //         fontSize: 10,
  //         color: "black",
  //       },
  //       subheader: {
  //         fontSize: 12,
  //         bold: true,
  //       },
  //     },
  //     defaultStyle: {
  //       font: "Roboto",
  //       fontSize: 8,
  //     },
  //   };

  //   return this.createPdfBuffer(documentDefinition);
  // }


  public async generateAttendanceReportPDFBuffer(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30",
    status: string = "Todos"
  ): Promise<Buffer> {

    const replacements: Record<string, any> = {
      startDate,
      endDate,
    };

    let havingClause = "";
    if (status !== "Todos" && status !== "") {
      havingClause = `HAVING "status" = :status`;
      replacements.status = status;
    }

    const dataQuery = `
      WITH RECURSIVE dates(d) AS (
          SELECT date(:startDate)
          UNION ALL
          SELECT date(d, '+1 day')
          FROM dates
          WHERE d < date(:endDate)
      )
      SELECT
          e.id,
          e.emp_code AS "Numero",
          e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
          dates.d AS "Fecha",
          MIN(time(p.punch_time)) AS "Entrada",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time))
              ELSE NULL
          END AS "Salida",
          dp.dept_name AS "Departamento",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN
                  ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24, 2)
              ELSE NULL
          END AS "TotalHorasRedondeadas",
          CASE strftime('%w', dates.d)
              WHEN '0' THEN 'Domingo'
              WHEN '1' THEN 'Lunes'
              WHEN '2' THEN 'Martes'
              WHEN '3' THEN 'Miércoles'
              WHEN '4' THEN 'Jueves'
              WHEN '5' THEN 'Viernes'
              WHEN '6' THEN 'Sábado'
          END AS "DiaSemana",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN 'Completo'
              WHEN COUNT(p.punch_time) = 1 THEN 'Pendiente'
              ELSE 'Sin Marcar'
          END AS "status"
      FROM hr_employee e
      CROSS JOIN dates
      LEFT JOIN att_punches p ON p.emp_id = e.id AND date(p.punch_time) = dates.d
      LEFT JOIN hr_department dp ON e.emp_dept = dp.id
      WHERE e.emp_active = 1
      GROUP BY dates.d, e.id, e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
      ${havingClause}
      ORDER BY dp.dept_name, e.emp_lastname, e.emp_firstname, dates.d ASC;
    `;

    const data = (await sequelize.query(dataQuery, {
      replacements,
      type: Sequelize.QueryTypes.SELECT,
    })) as Attendance[];

    const formattedData = data.map((item) => [
      { text: item.Nombre, alignment: "left" },
      {
        text: format(parseISO(item.Fecha), 'dd-MM-yyyy'),
        alignment: "center",
      },
      { text: item.Entrada, alignment: "center" },
      { text: item.Salida || "", alignment: "center" },
      { text: item.Departamento, alignment: "center" },
      {
        text: item.TotalHorasRedondeadas?.toString() || "",
        alignment: "center",
      },
      { text: item.DiaSemana, alignment: "center" },
      { text: item.status, alignment: "center" },
    ]);

    const documentDefinition: TDocumentDefinitions = {
      pageSize: "A4",
      pageMargins: [40, 60, 40, 60],
      header: (currentPage: number, pageCount: number) => {
        if (currentPage === 1) {
          return {
            columns: [
              {
                image: path.join(__dirname, "..", "assets/assets", "logo.png"),
                width: 70,
                margin: [40, 20, 0, 20],
              },
              {
                text: "Reporte de Asistencia",
                alignment: "center",
                fontSize: 18,
                bold: true,
                margin: [0, 25, 50, 0],
              },
            ],
          };
        } else {
          return null;
        }
      },
      footer: {
        columns: [
          {
            text: `Fecha de Reporte: ${new Date().toLocaleDateString("es-ES")}`,
            alignment: "left",
            margin: [40, 0, 0, 0],
          },
        ],
        margin: [0, 0, 0, 30],
      },
      content: [
        {
          table: {
            headerRows: 1,
            widths: ["25%", "15%", "10%", "10%", "10%", "10%", "10%", "10%"],
            body: [
              [
                { text: "Nombre", style: "tableHeader" },
                { text: "Fecha", style: "tableHeader" },
                { text: "Entrada", style: "tableHeader" },
                { text: "Salida", style: "tableHeader" },
                { text: "Departamento", style: "tableHeader" },
                { text: "Horas", style: "tableHeader" },
                { text: "Día", style: "tableHeader" },
                { text: "Estado", style: "tableHeader" },
              ],
              ...formattedData,
            ],
          },
          layout: {
            fillColor: (rowIndex: number) =>
              rowIndex === 0
                ? "#CCCCCC"
                : rowIndex % 2 === 0
                  ? "#F5F5F5"
                  : null,
            paddingLeft: () => 6,
            paddingRight: () => 6,
            paddingTop: () => 4,
            paddingBottom: () => 4,
          },
          margin: [0, 20, 0, 0],
        },
      ],
      styles: {
        tableHeader: {
          bold: true,
          fontSize: 10,
          color: "black",
        },
        subheader: {
          fontSize: 12,
          bold: true,
        },
      },
      defaultStyle: {
        font: "Roboto",
        fontSize: 8,
      },
    };

    return this.createPdfBuffer(documentDefinition);
  }

  public async generateAttendanceReportExcel(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30",
    status: string = "Todos"
  ): Promise<ExcelJS.Workbook> {
    const replacements: Record<string, any> = {
      startDate,
      endDate,
    };

    let havingClause = "";
    if (status !== "Todos" && status !== "") {
      havingClause = `HAVING "status" = :status`;
      replacements.status = status;
    }

    const dataQuery = `
      WITH RECURSIVE dates(d) AS (
        SELECT date(:startDate)
        UNION ALL
        SELECT date(d, '+1 day')
        FROM dates
        WHERE d < date(:endDate)
      )
      SELECT
        e.id,
        e.emp_code AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        dates.d AS "Fecha",
        MIN(time(p.punch_time)) AS "Entrada",
        CASE
          WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time))
          ELSE NULL
        END AS "Salida",
        dp.dept_name AS "Departamento",
        CASE
          WHEN COUNT(p.punch_time) > 1 THEN
            ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24, 2)
          ELSE NULL
        END AS "TotalHorasRedondeadas",
        CASE strftime('%w', dates.d)
          WHEN '0' THEN 'Domingo'
          WHEN '1' THEN 'Lunes'
          WHEN '2' THEN 'Martes'
          WHEN '3' THEN 'Miércoles'
          WHEN '4' THEN 'Jueves'
          WHEN '5' THEN 'Viernes'
          WHEN '6' THEN 'Sábado'
        END AS "DiaSemana",
        CASE
          WHEN COUNT(p.punch_time) > 1 THEN 'Completo'
          WHEN COUNT(p.punch_time) = 1 THEN 'Pendiente'
          ELSE 'Sin Marcar'
        END AS "status"
      FROM hr_employee e
      CROSS JOIN dates
      LEFT JOIN att_punches p ON p.emp_id = e.id AND date(p.punch_time) = dates.d
      LEFT JOIN hr_department dp ON e.emp_dept = dp.id
      WHERE e.emp_active = 1
      GROUP BY dates.d, e.id, e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
      ${havingClause}
      ORDER BY dp.dept_name, e.emp_lastname, e.emp_firstname, dates.d ASC;
    `;

    const data = (await sequelize.query(dataQuery, {
      replacements,
      type: Sequelize.QueryTypes.SELECT,
    })) as Attendance[];

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Reporte de Asistencia");

    worksheet.columns = [
      { header: "Nombre", key: "Nombre", width: 30 },
      { header: "Fecha", key: "Fecha", width: 15 },
      { header: "Entrada", key: "Entrada", width: 15 },
      { header: "Salida", key: "Salida", width: 15 },
      { header: "Departamento", key: "Departamento", width: 20 },
      { header: "Horas", key: "TotalHorasRedondeadas", width: 10 },
      { header: "Día", key: "DiaSemana", width: 10 },
      { header: "Estado", key: "status", width: 15 },
    ];

    data.forEach((record) => {
      worksheet.addRow({
        Nombre: record.Nombre,
        Fecha: format(parseISO(record.Fecha), 'dd/MM/yyyy'),
        Entrada: record.Entrada,
        Salida: record.Salida || "",
        Departamento: record.Departamento,
        TotalHorasRedondeadas: record.TotalHorasRedondeadas ?? "",
        DiaSemana: record.DiaSemana,
        status: record.status,
      });
    });

    worksheet.getRow(1).font = { bold: true };

    return workbook;
  }


  // generateControlReportExcel
  public async generateControlReportExcel(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30",
    department: string = "Todos"
  ): Promise<ExcelJS.Workbook> {
    const replacements: Record<string, any> = { startDate, endDate, department };
    const settingsFilePath = "./src/data/data/settings.json"; // Ruta del archivo settings.json
    const settingsData = fs.readFileSync(settingsFilePath, "utf8");
    const settings = JSON.parse(settingsData); // Leer y parsear settings.json


    let havingClause = "";
    if (department !== "Todos" && department !== "") {
      havingClause = `HAVING "DepartamentoID" = :department`;
      replacements.department = department;
    }

    const dataQuery = `
      WITH RECURSIVE dates(d) AS (
          SELECT date(:startDate)
          UNION ALL
          SELECT date(d, '+1 day')
          FROM dates
          WHERE d < date(:endDate)
      )
      SELECT
          e.id AS "EmpleadoID",
          e.emp_code AS "Numero",
          e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
          e.emp_firstname AS "Emp_firstname",
          e.emp_lastname AS "Emp_lastname",
          dp.id AS "DepartamentoID",
          dp.dept_name AS "Departamento",
          dates.d AS "Fecha",
          MIN(time(p.punch_time)) AS "Entrada",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time))
              ELSE NULL
          END AS "Salida",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN
                  ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24, 2)
              ELSE NULL
          END AS "TotalHorasRedondeadas",
          CASE strftime('%w', dates.d)
              WHEN '0' THEN 'Domingo'
              WHEN '1' THEN 'Lunes'
              WHEN '2' THEN 'Martes'
              WHEN '3' THEN 'Miércoles'
              WHEN '4' THEN 'Jueves'
              WHEN '5' THEN 'Viernes'
              WHEN '6' THEN 'Sábado'
          END AS "DiaSemana",
          CASE
              WHEN COUNT(p.punch_time) > 1 THEN 'Completo'
              WHEN COUNT(p.punch_time) = 1 THEN 'Pendiente'
              ELSE 'Sin Marcar'
          END AS "status",
          CASE
              WHEN COUNT(p.punch_time) = 0 THEN 'Z'
              ELSE NULL
          END AS "TipoZ",
          CASE
              WHEN COUNT(p.punch_time) = 1 AND strftime('%H:%M', MIN(p.punch_time)) < '12:00' THEN 'HI'
              ELSE NULL
          END AS "HI",
          CASE
              WHEN COUNT(p.punch_time) = 1 AND strftime('%H:%M', MIN(p.punch_time)) >= '12:00' THEN 'HS'
              ELSE NULL
          END AS "HS",
          CASE
              WHEN ea.paycode_id = 12 THEN 'Si'
              ELSE 'No'
          END AS "Vacaciones",
          CASE
              WHEN ea.paycode_id = 13 THEN 'Si'
              ELSE 'No'
          END AS "Permiso",
          CASE
              WHEN ea.paycode_id = 11 THEN 'Si'
              ELSE 'No'
          END AS "Enfermedad"
      FROM
          hr_employee e
              CROSS JOIN dates
              LEFT JOIN att_punches p ON p.emp_id = e.id AND date(p.punch_time) = dates.d
              LEFT JOIN hr_department dp ON e.emp_dept = dp.id
              LEFT JOIN att_exceptionassign ea ON ea.employee_id = e.id
      WHERE
          e.emp_active = 1
      GROUP BY
          dates.d,
          e.id,
          e.emp_code,
          e.emp_firstname,
          e.emp_lastname,
          dp.dept_name,
          ea.paycode_id
      ${havingClause}
      ORDER BY
          dp.dept_name,
          e.emp_lastname,
          e.emp_firstname,
          dates.d ASC;
    `;

    const data = (await sequelize.query(dataQuery, {
      replacements,
      type: Sequelize.QueryTypes.SELECT,
    })) as any[];

    console.log(settings)
    // console.log(data)

    // Extract unique dates
    const uniqueDates = [...new Set(data.map(record => record.Fecha))].sort();
    const headers = ["N°", "M", "NOMBRES", ...uniqueDates.flatMap(date => [date, "HE", "HS", "NOVEDADES"])];

    // Agrupar los datos por usuario
    const groupedData: Record<string, any> = {};

    data.forEach(record => {
      const userKey = record.Nombre;
      if (!groupedData[userKey]) {
        groupedData[userKey] = {
          "N°": record.EmpleadoID,
          "M": "",
          "NOMBRES": record.Nombre,
          dates: {}
        };
      }


      // >>>>>>>>>>>>>>>>>>>>>>>>>>>>

      // Validar y asignar el campo Novedades
      let novedades = "";

      if (Number(record.PaycodeID) === 11) {
        // Enfermedad
        novedades = settings.sickLeaveSymbol;
      } else if (Number(record.PaycodeID) === 12) {
        // Vacaciones
        novedades = settings.vacationSymbol;
      } else if (record.HI === "HI") {
        // Solo entrada
        novedades = settings.entranceOnlySymbol;
      } else if (record.Permiso === "Si") {
        // Permiso
        novedades = settings.earlyExitSymbol;
      } else if (record.HS === "HS") {
        // Solo salida
        novedades = settings.exitOnlySymbol;
      } else if (record.TipoZ === "Z") {
        // Ausencia
        novedades = settings.absenceSymbol;
      } else {
        // Total de horas o símbolo predeterminado
        // novedades =
        //   record.TotalHorasRedondeadas !== null &&
        //     record.TotalHorasRedondeadas !== 0
        //     ? record.TotalHorasRedondeadas.toString()
        //     : settings.entranceOnlySymbol;
      }

      //>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


      groupedData[userKey].dates[record.Fecha] = {
        Entrada: record.Entrada || "",
        Salida: record.Salida || "",
        Novedades: novedades || "", // Puedes ajustar si tienes datos reales de novedades
      };
    });

    // =============================================

    // Crear el workbook y la hoja de trabajo
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Reporte de Asistencia");

    // =============================================

    // Agregar dos filas vacías al inicio y ajustar su altura
    const emptyRow1 = worksheet.addRow([]);
    emptyRow1.height = 30; // Establecer la altura de la primera fila vacía

    const emptyRow2 = worksheet.addRow([]);
    emptyRow2.height = 30; // Establecer la altura de la segunda fila vacía

    // Unir las celdas desde A1 hasta K2
    worksheet.mergeCells('A1:K2');

    // Cargar la imagen al workbook
    const imageHeH = workbook.addImage({
      filename: './src/assets/HeH.png', // Ruta de la imagen
      extension: 'png',
    });

    // Añadir la imagen flotando en una posición específica
    worksheet.addImage(imageHeH, {
      tl: { col: 0.2, row: 0.3 }, // Top-left: esquina superior izquierda (columna A, fila 1)
      ext: { width: 100, height: 60 }, // Tamaño de la imagen en píxeles
    });

    // Cargar la imagen al workbook
    const imageMM = workbook.addImage({
      filename: './src/assets/MM.png', // Ruta de la imagen
      extension: 'png',
    });

    // Añadir la imagen flotando en una posición específica
    worksheet.addImage(imageMM, {
      tl: { col: 16.3, row: 0 }, // Top-left: esquina superior izquierda (columna A, fila 1)
      ext: { width: 140, height: 80 }, // Tamaño de la imagen en píxeles
    });

    // =============================================

    // Definir columnas con anchos personalizados
    worksheet.columns = [
      { key: "N°", width: 5 },
      { key: "M", width: 5 },
      { key: "NOMBRES", width: 30 },
      ...uniqueDates.flatMap((date) => [
        { key: `${date}_NOVEDADES`, width: 15 },
      ]),
    ];

    // Agregar los headers
    const headerRow = worksheet.getRow(3);
    headerRow.values = headers;

    // Aumentar la altura de las cabeceras
    headerRow.height = 25;

    // Aplicar estilos a los headers y mergear las fechas (solo HE y HS)
    uniqueDates.forEach((date, index) => {
      const baseCol = 4 + (index * 3); // Avanzar 3 columnas por fecha
      worksheet.mergeCells(3, baseCol, 3, baseCol + 1); // Merge solo para HE y HS
      const cell = worksheet.getCell(3, baseCol);
      cell.value = date; // Asignar la fecha como encabezado
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' },
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      // Encabezado independiente para NOVEDADES
      const novedadesCell = worksheet.getCell(3, baseCol + 2);
      novedadesCell.value = "NOVEDADES";
      novedadesCell.font = { bold: true };
      novedadesCell.alignment = { horizontal: "center", vertical: "middle" };
      novedadesCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' },
      };
      novedadesCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Aplicar el mismo estilo a las cabeceras estáticas: N°, M, NOMBRES
    [1, 2].forEach((col) => {
      const cell = worksheet.getCell(3, col);
      cell.font = { bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }, // Color gris claro (igual al de las fechas)
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Agregar las filas de datos
    let rowCounter = 4;
    Object.values(groupedData).forEach((userData, index) => {
      const dataRowTop = worksheet.getRow(rowCounter);
      const dataRowBottom = worksheet.getRow(rowCounter + 1);

      // N°, M, y NOMBRES
      dataRowTop.getCell(1).value = index + 1;
      dataRowTop.getCell(2).value = userData.M;
      dataRowTop.getCell(3).value = userData.NOMBRES;

      // Merge cells for N°, M, and NOMBRES
      worksheet.mergeCells(rowCounter, 1, rowCounter + 1, 1);
      worksheet.mergeCells(rowCounter, 2, rowCounter + 1, 2);
      worksheet.mergeCells(rowCounter, 3, rowCounter + 1, 3);

      // Añadir borde y color alternado a las columnas estáticas
      const staticColumns = [1, 2, 3];
      const backgroundColor = index % 2 === 0 ? 'FFEFEFEF' : 'FFFFFFFF'; // Gris claro alternado
      staticColumns.forEach((col) => {
        const cellTop = dataRowTop.getCell(col);
        const cellBottom = dataRowBottom.getCell(col);

        // Establecer fondo y bordes
        [cellTop, cellBottom].forEach((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: backgroundColor },
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      // Fechas y NOVEDADES
      uniqueDates.forEach((date, dateIndex) => {
        const baseCol = 4 + (dateIndex * 3); // Avanzar 3 columnas por fecha
        const dateData = userData.dates[date] || { Entrada: "", Salida: "", Novedades: "" };

        // Primera columna: HE
        const heCell = dataRowTop.getCell(baseCol);
        heCell.value = "HE:";
        heCell.font = { size: 8 };
        heCell.alignment = { vertical: "middle", horizontal: "center" };
        heCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: undefined, // Quitar borde derecho
        };
        heCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } };

        // Segunda columna: HS
        const hsCell = dataRowBottom.getCell(baseCol);
        hsCell.value = "HS:";
        hsCell.font = { size: 8 };
        hsCell.alignment = { vertical: "middle", horizontal: "center" };
        hsCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: undefined, // Quitar borde derecho
        };
        hsCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } };

        // Segunda columna: horarios (Entrada y Salida)
        const entradaCell = dataRowTop.getCell(baseCol + 1);
        entradaCell.value = dateData.Entrada || ""; // Mostrar vacío si no hay dato
        entradaCell.border = {
          top: { style: "thin" },
          left: undefined, // Quitar borde izquierdo
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        entradaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } };

        const salidaCell = dataRowBottom.getCell(baseCol + 1);
        salidaCell.value = dateData.Salida || ""; // Mostrar vacío si no hay dato
        salidaCell.border = {
          top: { style: "thin" },
          left: undefined, // Quitar borde izquierdo
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        salidaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } };

        // Tercera columna: NOVEDADES
        const novedadesCell = worksheet.getCell(rowCounter, baseCol + 2);
        novedadesCell.value = dateData.Novedades || ""; // Mostrar vacío si no hay dato
        novedadesCell.alignment = { horizontal: "center", vertical: "middle" };
        worksheet.mergeCells(rowCounter, baseCol + 2, rowCounter + 1, baseCol + 2);
        novedadesCell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        novedadesCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: backgroundColor } };
      });

      // Reducir el ancho de las columnas de HE y HS
      for (let colIndex = 4; colIndex <= worksheet.columnCount; colIndex += 3) {
        worksheet.getColumn(colIndex).width = 3; // Establecer el ancho de HE
      }
      for (let colIndex = 5; colIndex <= worksheet.columnCount; colIndex += 3) {
        worksheet.getColumn(colIndex).width = 12; // Establecer el ancho de Entrada/Salida
      }

      // Asegurar ancho uniforme para columnas NOVEDADES
      uniqueDates.forEach((date, dateIndex) => {
        const novedadesCol = worksheet.getColumn(4 + (dateIndex * 3) + 2); // Índice de la columna NOVEDADES
        novedadesCol.width = 15; // Mantener ancho uniforme
      });

      dataRowTop.commit();
      dataRowBottom.commit();
      rowCounter += 2;
    });


    // Aplicar alineación específica para la columna "NOMBRES"
    worksheet.getColumn(1).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getColumn(2).alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getColumn(3).alignment = { horizontal: "left", vertical: "middle" };

    // Centrar solo la cabecera de NOMBRES
    const nombresHeaderCell = worksheet.getCell(3, 3);
    nombresHeaderCell.font = { bold: true };
    nombresHeaderCell.alignment = { horizontal: "center", vertical: "middle" };
    nombresHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' },
    };
    nombresHeaderCell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // Ajustar el rango de columnas activas
    worksheet.spliceColumns(3 + uniqueDates.length * 3 + 1, worksheet.columnCount - (3 + uniqueDates.length * 3));

    // Congelar las primeras tres filas y las primeras tres columnas
    worksheet.views = [
      { state: 'frozen', xSplit: 3, ySplit: 3, topLeftCell: 'D4', activeCell: 'D4' },
    ];

    // =============================================

    // Configurar el texto de la celda y centrarlo
    const titleCell = worksheet.getCell('A1');
    titleCell.value = "CONTROL DE INGRESO Y SALIDA DE PERSONAL";
    titleCell.font = { bold: true, size: 18 }; // Negrita y tamaño grande
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }; // Centrado horizontal y verticalmente

    // =============================================

    // Agregar "Codigo:" en la columna L, fila A
    worksheet.getCell('L1').value = "Código:";
    worksheet.getCell('L1').alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell('L1').font = { bold: true };

    // Agregar EL CODIGO en la columna T, fila A
    worksheet.getCell('M1').value = "F-RUH-SGF-001";
    worksheet.getCell('M1').alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell('M1').font = { bold: true };

    // Agregar "Versión:" en la columna L, fila B
    worksheet.getCell('L2').value = "Versión:";
    worksheet.getCell('L2').alignment = { horizontal: "center", vertical: "middle" };
    worksheet.getCell('L2').font = { bold: true };

    // Agregar LA VERSION en la columna R, fila B
    worksheet.getCell('M2').value = "04";
    worksheet.getCell('M2').alignment = { horizontal: "center", vertical: "middle" };


    // Agregar bordes y formato a las celdas unidas L1:O1 y L2:O2
    ['L1:O1', 'L2:O2'].forEach(range => {
      const mergedCell = worksheet.getCell(range.split(':')[0]);
      mergedCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Unir las columnas M a O en las filas A y B (en blanco)
    worksheet.mergeCells('M1:O1');
    worksheet.mergeCells('M2:O2');

    // Agregar bordes y formato a las celdas unidas M1:O1 y M2:O2
    ['M1:O1', 'M2:O2'].forEach(range => {
      const mergedCell = worksheet.getCell(range.split(':')[0]);
      mergedCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Unir las columnas P a R en las filas A y B
    worksheet.mergeCells('P1:R2');

    // Estilizar y agregar bordes a las celdas unidas P1:R2
    const mergedCellPR = worksheet.getCell('P1');
    mergedCellPR.value = ""; // Dejar en blanco (puedes agregar texto si es necesario)
    mergedCellPR.alignment = { horizontal: "center", vertical: "middle" };
    mergedCellPR.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // =============================================
    // Tabla estatica
    // =============================================

    // Añadir tabla estática después de lo generado dinámicamente
    const lastDynamicRow = worksheet.lastRow?.number || worksheet.rowCount;
    const startStaticTableRow = lastDynamicRow + 2;

    // Ajustar altura de la fila
    worksheet.getRow(startStaticTableRow).height = 25;

    //

    // Primera fila de la tabla estática (REALIZADO POR)
    worksheet.mergeCells(`A${startStaticTableRow}:L${startStaticTableRow}`);
    const headerCell = worksheet.getCell(`A${startStaticTableRow}`);
    headerCell.value = "REALIZADO POR";
    headerCell.font = { bold: true };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' }, };

    // Definir las divisiones
    worksheet.mergeCells(`M${startStaticTableRow}:O${startStaticTableRow}`);
    worksheet.mergeCells(`P${startStaticTableRow}:R${startStaticTableRow}`);

    // Agregar el texto "REVISADO POR:" en la primera columna de la segunda fila
    const header_revisadoCell = worksheet.getCell(`M${startStaticTableRow}`);
    header_revisadoCell.value = "REVISADO POR";
    header_revisadoCell.font = { bold: true };
    header_revisadoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    header_revisadoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' }, };


    // Agregar el texto "APROBADO POR:" en la primera columna de la segunda fila
    const header_aprobadoCell = worksheet.getCell(`P${startStaticTableRow}`);
    header_aprobadoCell.value = "APROBADO POR";
    header_aprobadoCell.font = { bold: true };
    header_aprobadoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    header_aprobadoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' }, };

    //

    // Segunda fila de la tabla estática (dividida en 4 secciones)
    const secondRow = startStaticTableRow + 1;

    // Ajustar altura de la fila
    worksheet.getRow(secondRow).height = 20;

    // Definir las divisiones
    worksheet.mergeCells(`A${secondRow}:C${secondRow}`);
    worksheet.mergeCells(`D${secondRow}:F${secondRow}`);
    worksheet.mergeCells(`G${secondRow}:I${secondRow}`);
    worksheet.mergeCells(`J${secondRow}:L${secondRow}`);
    worksheet.mergeCells(`M${secondRow}:O${secondRow}`);
    worksheet.mergeCells(`P${secondRow}:R${secondRow}`);

    // Agregar el texto "FIRMA:" en la primera columna de la segunda fila
    const firmaCell = worksheet.getCell(`A${secondRow}`);
    firmaCell.value = "FIRMA:";
    firmaCell.font = { bold: true };
    firmaCell.alignment = { horizontal: 'center', vertical: 'middle' };
    firmaCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    // Agregar estilos a la celda de firma
    const firma_revisadoCell = worksheet.getCell(`M${secondRow}`);
    firma_revisadoCell.font = { bold: true };
    firma_revisadoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    firma_revisadoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    // Agregar estilos
    const firma_aprobadoCell = worksheet.getCell(`P${secondRow}`);
    firma_aprobadoCell.font = { bold: true };
    firma_aprobadoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    firma_aprobadoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    //

    // Tercera fila de la tabla estática (dividida en 4 secciones)
    const threeRow = secondRow + 1;

    // Ajustar altura de la fila
    worksheet.getRow(threeRow).height = 20;

    // Definir las divisiones
    worksheet.mergeCells(`A${threeRow}:C${threeRow}`);
    worksheet.mergeCells(`D${threeRow}:F${threeRow}`);
    worksheet.mergeCells(`G${threeRow}:I${threeRow}`);
    worksheet.mergeCells(`J${threeRow}:L${threeRow}`);
    worksheet.mergeCells(`M${threeRow}:O${threeRow}`);
    worksheet.mergeCells(`P${threeRow}:R${threeRow}`);

    // Agregar el texto "NOMBRE:" en la primera columna de la segunda fila
    const nameCell = worksheet.getCell(`A${threeRow}`);
    nameCell.value = "NOMBRE:";
    nameCell.font = { bold: true };
    nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
    nameCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    const name_revisadoCell = worksheet.getCell(`M${threeRow}`);
    name_revisadoCell.value = "Ing. Jessenia Maldonado"; // Asegúrate de que no haya espacios innecesarios antes del texto
    name_revisadoCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; // wrapText habilitado

    const name_aprovadoCell = worksheet.getCell(`P${threeRow}`);
    name_aprovadoCell.value = "Ing. Iván Ontaneda"; // Asegúrate de que no haya espacios innecesarios antes del texto
    name_aprovadoCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; // wrapText habilitado

    // 

    // Cuarta fila de la tabla estática (dividida en 4 secciones)
    const fourRow = threeRow + 1;

    // Ajustar altura de la fila
    worksheet.getRow(fourRow).height = 20;

    // Definir las divisiones
    worksheet.mergeCells(`A${fourRow}:C${fourRow}`);
    worksheet.mergeCells(`D${fourRow}:F${fourRow}`);
    worksheet.mergeCells(`G${fourRow}:I${fourRow}`);
    worksheet.mergeCells(`J${fourRow}:L${fourRow}`);
    worksheet.mergeCells(`M${fourRow}:O${fourRow + 2}`);
    worksheet.mergeCells(`P${fourRow}:R${fourRow + 2}`);

    // Agregar el texto "PERIODO:" en la primera columna de la segunda fila
    const periodCell = worksheet.getCell(`A${fourRow}`);
    periodCell.value = "PERIODO: (Fecha)";
    periodCell.font = { bold: true };
    periodCell.alignment = { horizontal: 'center', vertical: 'middle' };
    periodCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    // Agregar estilos a la celda de firma
    const period_revisadoCell = worksheet.getCell(`M${fourRow}`);
    period_revisadoCell.value = "Jefe Administrativo \nIndustria Metal Mecánica"; // Asegúrate de que no haya espacios innecesarios antes del texto
    period_revisadoCell.font = { bold: true };
    period_revisadoCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; // wrapText habilitado

    // Agregar estilos
    const period_aprobadoCell = worksheet.getCell(`P${fourRow}`);
    period_aprobadoCell.value = "Superintendente \nIndustria Metal Mecánica"; // Asegúrate de que no haya espacios innecesarios antes del texto
    period_aprobadoCell.font = { bold: true };
    period_aprobadoCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; // wrapText habilitado

    // 

    // Quinta fila de la tabla estática (dividida en 4 secciones)
    const fiveRow = fourRow + 1;

    // Ajustar altura de la fila
    worksheet.getRow(fiveRow).height = 20;

    // Definir las divisiones
    worksheet.mergeCells(`A${fiveRow}:C${fiveRow}`);
    worksheet.mergeCells(`D${fiveRow}:F${fiveRow}`);
    worksheet.mergeCells(`G${fiveRow}:I${fiveRow}`);
    worksheet.mergeCells(`J${fiveRow}:L${fiveRow}`);

    // Agregar el texto "HORARIO:" en la primera columna de la segunda fila
    const horarioCell = worksheet.getCell(`A${fiveRow}`);
    horarioCell.value = "HORARIO:";
    horarioCell.font = { bold: true };
    horarioCell.alignment = { horizontal: 'center', vertical: 'middle' };
    horarioCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' }, };

    // Ultima fila de la tabla estática
    const sixRow = fiveRow + 1;

    // Sexta fila de la tabla estática (Guardia)
    worksheet.mergeCells(`A${sixRow}:L${sixRow}`);
    const footerCell = worksheet.getCell(`A${sixRow}`);
    footerCell.value = "Guardia";
    footerCell.font = { bold: true };
    footerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    footerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CCCCCC' }, };

    //

    // Aplicar bordes a todas las celdas de la tabla estática
    for (let row = startStaticTableRow; row <= sixRow; row++) {
      for (let col = 1; col <= 18; col++) { // De la columna A (1) a la columna R (18)
        const cell = worksheet.getCell(row, col);
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }

    // =============================================

    // Configurar ajustes de impresión
    worksheet.pageSetup = {
      paperSize: 9, // Tamaño de papel A4
      orientation: 'landscape', // Orientación horizontal
      fitToPage: true, // Ajustar al tamaño de la página
      fitToWidth: 0, // Permitir varias páginas en ancho
      fitToHeight: 1, // Ajustar todas las filas a una página
      margins: { // Márgenes en pulgadas
        left: 0.5,
        right: 0.5,
        top: 0.5,
        bottom: 0.5,
        header: 0.3,
        footer: 0.3,
      },
    };

    return workbook;

  }

}
