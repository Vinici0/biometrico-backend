import { Request, Response, RequestHandler } from "express";
import { ExceptionService } from "../../services/exception.service";
import ExcelJS from "exceljs";

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
}
