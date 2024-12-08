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
    const { startDate, endDate, page, pageSize } = req.body;
    
    this.categoryService
      .getMonthlyAttendanceReport(startDate, endDate, page, pageSize)
      .then((data) => {
        return HttpResponseHandler.success(res, data);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  searchAttendance: RequestHandler = (req: Request, res: Response) => {
    const { name, page, pageSize, startDate, endDate, department, status, employeeId } = req.body;
    console.log(startDate, endDate, page, pageSize);

    this.categoryService
      .searchAttendance({ name, page, pageSize, startDate, endDate, department, status, employeeId })
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

    const { startDate , endDate  } = req.query;
    
    try {
      const buffer = await this.categoryService.downloadMonthlyAttendanceReport(
        {
          startDate: startDate as string,
          endDate: endDate as string, 
        }
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
