import { Request, Response, RequestHandler } from "express";
import { ExceptionService } from "../../services/exception.service";
import ExcelJS from "exceljs";

const dataResult = [
  {
    id: 16,
    Numero: "",
    Nombre: "JOSE LINO PINCAY",
    Fecha: "2022-05-26",
    Entrada: "00:05:22",
    Salida: "06:02:37",
    Departamento: "Personal.HeH",
    TotalHorasRedondeadas: 6,
    DiaSemana: "Jueves",
    status: "Completo",
  },
  {
    id: 45,
    Numero: "",
    Nombre: "JOSE ZABALA GABIDIA",
    Fecha: "2022-05-26",
    Entrada: "11:40:26",
    Salida: "12:31:36",
    Departamento: "Personal.HeH",
    TotalHorasRedondeadas: 1,
    DiaSemana: "Jueves",
    status: "Completo",
  },
  {
    id: 10,
    Numero: "",
    Nombre: "LIMBERG COBOS CHOES",
    Fecha: "2022-05-26",
    Entrada: "15:11:02",
    Salida: null,
    Departamento: "Personal.HeH",
    TotalHorasRedondeadas: null,
    DiaSemana: "Jueves",
    status: "Pendiente",
  },
  {
    id: 41,
    Numero: "",
    Nombre: "MIGUEL TOALA SEGUICHE",
    Fecha: "2022-05-26",
    Entrada: "01:00:41",
    Salida: "23:00:21",
    Departamento: "Personal.HeH",
    TotalHorasRedondeadas: 22,
    DiaSemana: "Jueves",
    status: "Completo",
  },
];
export class ExceptionController {
  constructor(private exceptionService: ExceptionService) {}

  public getExceptionReport: RequestHandler = async (
    req: Request,
    res: Response
  ) => {
    const { startDate, endDate } = req.query;
    console.log(`Generando reporte desde ${startDate} hasta ${endDate}`);

    try {
      const pdfBuffer = await this.exceptionService.getExceptionReport(
        (startDate as string) || "",
        (endDate as string) || ""
      );

      // Configurar las cabeceras de la respuesta para el PDF
      res.set({
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename=exception-report-${Date.now()}.pdf`,
        "Content-Length": pdfBuffer.length,
      });

      // Enviar el buffer como respuesta
      res.send(pdfBuffer);
    } catch (error) {
      console.error("Error generando el reporte de excepciones:", error);
      res.status(500).json({ message: "Error generando el reporte." });
    }
  };

  public getExceptionReportExcel: RequestHandler = async (
    req: Request,
    res: Response
  ) => {
    const { startDate, endDate } = req.query;
    console.log(`Generando reporte Excel desde ${startDate} hasta ${endDate}`);

    try {
      const workbook = await this.exceptionService.getExceptionReportExcel(
        (startDate as string) || "",
        (endDate as string) || ""
      );

      // Convertir el workbook a un buffer
      const buffer = await workbook.xlsx.writeBuffer();

      // Configurar las cabeceras de la respuesta para el Excel
      res.set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename=exception-report-${Date.now()}.xlsx`,
        "Content-Length": Buffer.byteLength(buffer),
      });

      // Enviar el buffer como respuesta
      res.send(buffer);
    } catch (error) {
      console.error("Error generando el reporte de Excel:", error);
      res.status(500).json({ message: "Error generando el reporte de Excel." });
    }
  };

  public generateAttendanceReportPDFBuffer: RequestHandler = async (
    req: Request,
    res: Response
  ) => {
    const { startDate, endDate } = req.query;
    console.log(`Generando reporte desde ${startDate} hasta ${endDate}`);

    try {
      const pdfBuffer =
        await this.exceptionService.generateAttendanceReportPDFBuffer(
          startDate as string,
          endDate as string
        );

      // Configurar las cabeceras de la respuesta para el PDF
      res.set({
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename=exception-report-${Date.now()}.pdf`,
        "Content-Length": pdfBuffer.length,
      });

      // Enviar el buffer como respuesta
      res.send(pdfBuffer);
    } catch (error) {
      console.error("Error generando el reporte de excepciones:", error);
      res.status(500).json({ message: "Error generando el reporte." });
    }
  };

  // generateAttendanceReportExcel
  public generateAttendanceReportExcel: RequestHandler = async (
    req: Request,
    res: Response
  ) => {
    const { startDate, endDate } = req.query;

    try {
      const workbook =
        await this.exceptionService.generateAttendanceReportExcel(
          startDate as string,
          endDate as string
        );

      // Convertir el workbook a un buffer
      const buffer = await workbook.xlsx.writeBuffer();

      // Configurar las cabeceras de la respuesta para el Excel
      res.set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename=exception-report-${Date.now()}.xlsx`,
        "Content-Length": Buffer.byteLength(buffer),
      });

      // Enviar el buffer como respuesta

      res.send(buffer);
    } catch (error) {
      console.error("Error generando el reporte de excepciones:", error);
      res.status(500).json({ message: "Error generando el reporte." });
    }
  };
}
