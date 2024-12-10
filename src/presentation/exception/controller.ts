import { Request, Response, RequestHandler } from "express";
import { ExceptionService } from "../../services/exception.service";

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
}
