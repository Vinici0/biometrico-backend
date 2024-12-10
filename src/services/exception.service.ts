// src/services/exception.service.ts

import PdfPrinter from "pdfmake";
import path from "path";
import Sequelize from "sequelize";
import sequelize from "../config/database";
import { TDocumentDefinitions } from "pdfmake/interfaces";
import { resultsInit } from "../data/init";

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

export class ExceptionService {
  private printer: PdfPrinter;

  constructor() {
    // Definir las fuentes
    this.printer = new PdfPrinter({
      Roboto: {
        normal: path.join(__dirname, "../fonts/Roboto-Regular.ttf"),
        bold: path.join(__dirname, "../fonts/Roboto-Medium.ttf"),
        italics: path.join(__dirname, "../fonts/Roboto-Italic.ttf"),
        bolditalics: path.join(__dirname, "../fonts/Roboto-MediumItalic.ttf"),
      },
    });
  }

  public async getExceptionReport(
    startDate: string = "2024-09-01",
    endDate: string = "2024-09-30"
  ): Promise<Buffer> {
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
}
