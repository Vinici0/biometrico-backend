import { Request, Response, RequestHandler } from "express";
import { AttendanceService } from "../../services/attendance.service";
import { HttpResponseHandler } from "../../domain/response/http-response-handler";
import { CustomError } from "../../domain/response/custom.error";

export class AttendanceController {
  constructor(public readonly categoryService: AttendanceService) {}

  getMonthlyAttendanceReport: RequestHandler = (
    req: Request,
    res: Response
  ) => {
    const { startDate, endDate } = req.body;

    this.categoryService
      .getMonthlyAttendanceReport(startDate, endDate)
      .then((data) => {
        return HttpResponseHandler.success(res, data);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  //descargar reporte de asistencia xml
  downloadMonthlyAttendanceReport: RequestHandler = async (
    req: Request,
    res: Response
  ) => {
    // const { startDate = "", endDate = "" } = req.query; // Cambiar a `req.query` si estás enviando parámetros en la URL

    try {
      const buffer = await this.categoryService.downloadMonthlyAttendanceReport(
        {}
      );

      // Configura los encabezados de la respuesta para forzar la descarga
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="ReporteAsistencia_Firma.xlsx"`
      );

      // Envía el archivo como un adjunto en la respuesta
      res.send(buffer);
    } catch (error) {
      this.handleError(error, res);
    }
  };

  private handleError = (error: unknown, res: Response) => {
    if (error instanceof CustomError) {
      return HttpResponseHandler.error(
        res,
        error,
        error.message,
        error.statusCode
      );
    }

    console.log(`${error}`);
    return HttpResponseHandler.error(res, error);
  };
}
