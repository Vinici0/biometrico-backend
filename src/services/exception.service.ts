// src/services/exception.service.ts

import PdfPrinter from "pdfmake";
import path from "path";
import Sequelize from "sequelize";
import sequelize from "../config/database";
import { TDocumentDefinitions } from "pdfmake/interfaces";
import { resultsInit } from "../data/init";
import ExcelJS from "exceljs";

export interface ExceptionRecord {
  emp_pin: number;
  emp_firstname: string;
  emp_lastname: string;
  exception_date: string;
  starttime: string;
  endtime: string;
  total_horas_excepcion: number;
  tipo_de_excepcion: string;
}

export interface Attendance {
  id: number;
  Numero: string;
  Nombre: string;
  Fecha: string;
  Entrada: string;
  Salida: null | string;
  Departamento: string;
  TotalHorasRedondeadas: number | null;
  DiaSemana: string;
  status: string;
}

export class ExceptionService {
  private printer: PdfPrinter;

  constructor() {
    // Definir la ruta relativa a la carpeta de fuentes
    const fontsPath = path.join(__dirname, "..", "fonts");

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
                image: path.join(__dirname, "../assets/logo.png"),
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
                image: path.join(__dirname, "..", "assets", "logo.png"),
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
    endDate: string = "2024-09-30"
  ): Promise<Buffer> {

    // Consulta SQL para obtener los datos de asistencia
    const dataQuery = `
      SELECT 
        e.id, 
        e.emp_code AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        date(p.punch_time) AS "Fecha",
        MIN(time(p.punch_time)) AS "Entrada",
        CASE 
          WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time)) 
          ELSE NULL 
        END AS "Salida",
        dp.dept_name AS "Departamento",
        CASE 
          WHEN COUNT(p.punch_time) > 1 THEN ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24) 
          ELSE NULL 
        END AS "TotalHorasRedondeadas",
        CASE strftime('%w', date(p.punch_time))
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
          ELSE 'Pendiente'
        END AS "status"
      FROM 
        att_punches p
      JOIN 
        hr_employee e ON p.emp_id = e.id
      LEFT JOIN 
        hr_department dp ON e.emp_dept = dp.id
      WHERE 
        date(p.punch_time) BETWEEN :startDate AND :endDate
      GROUP BY 
        date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
      ORDER BY 
        date(p.punch_time)
    `;

    // Ejecutar la consulta SQL
    const data = (await sequelize.query(dataQuery, {
      replacements: { startDate, endDate },
      type: Sequelize.QueryTypes.SELECT,
    })) as Attendance[];

    const formattedData = data.map((item) => [
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
                image: path.join(__dirname, "..", "assets", "logo.png"),
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
            widths: ["25%", "10%", "10%", "15%", "10%", "10%", "10%", "10%"],
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
    endDate: string = "2024-09-30"
  ): Promise<ExcelJS.Workbook> {
    // Consulta SQL para obtener los datos de asistencia
    const dataQuery = `
      SELECT 
        e.id, 
        e.emp_code AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        date(p.punch_time) AS "Fecha",
        MIN(time(p.punch_time)) AS "Entrada",
        CASE 
          WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time)) 
          ELSE NULL 
        END AS "Salida",
        dp.dept_name AS "Departamento",
        CASE 
          WHEN COUNT(p.punch_time) > 1 THEN ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24) 
          ELSE NULL 
        END AS "TotalHorasRedondeadas",
        CASE strftime('%w', date(p.punch_time))
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
          ELSE 'Pendiente'
        END AS "status"
      FROM 
        att_punches p
      JOIN 
        hr_employee e ON p.emp_id = e.id
      LEFT JOIN 
        hr_department dp ON e.emp_dept = dp.id
      WHERE 
        date(p.punch_time) BETWEEN :startDate AND :endDate
      GROUP BY 
        date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
      ORDER BY 
        date(p.punch_time)
    `;

    // Ejecutar la consulta SQL
    const data = (await sequelize.query(dataQuery, {
      replacements: { startDate, endDate },
      type: Sequelize.QueryTypes.SELECT,
    })) as Attendance[];

    // Crear el workbook y el worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Reporte de Asistencia");

    // Definir las columnas del Excel
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

    // Agregar los datos al worksheet
    data.forEach((record) => {
      worksheet.addRow({
        Nombre: record.Nombre,
        Fecha: new Date(record.Fecha).toLocaleDateString("es-ES"),
        Entrada: record.Entrada,
        Salida: record.Salida || "",
        Departamento: record.Departamento,
        TotalHorasRedondeadas: record.TotalHorasRedondeadas ?? "",
        DiaSemana: record.DiaSemana,
        status: record.status,
      });
    });

    // Encabezados en negrita
    worksheet.getRow(1).font = { bold: true };

    return workbook;
  }
}
