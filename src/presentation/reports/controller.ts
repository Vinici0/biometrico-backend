import { Request, Response, RequestHandler } from "express";
import { AttendanceService } from "../../services/attendance.service";
import { HttpResponseHandler } from "../../domain/response/http-response-handler";
import { CustomError } from "../../domain/response/custom.error";
import { format } from "date-fns";

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
    const {
      name,
      page,
      pageSize,
      startDate,
      endDate,
      department,
      status,
      employeeId,
    } = req.body;
    console.log(startDate, endDate, page, pageSize);

    this.categoryService
      .searchAttendance({
        name,
        page,
        pageSize,
        startDate,
        endDate,
        department,
        status,
        employeeId,
      })
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
    const { startDate, endDate, department } = req.query;
    console.log(
      `Generando reporte desde ${startDate} hasta ${endDate} ${department}`
    );

    try {
      const buffer = await this.categoryService.downloadMonthlyAttendanceReport(
        {
          startDate: startDate as string,
          endDate: endDate as string,
          department: (department as string | undefined) || null,
        }
      );

      const currentDate = format(new Date(), "yyyy-MM-dd");

      // Configura los encabezados de la respuesta para forzar la descarga
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="ReporteAsistencia_${currentDate}.xlsx"`
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

  // =============================================================
  // Conteo total de dashboard

  totalDashboardReport: RequestHandler = (req: Request, res: Response) => {
    this.categoryService
      .totalDashboardReport()
      .then((data) => {
        return HttpResponseHandler.success(res, data);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  // Total de asistencias, aucencias y atrasos en los ultimos 6 meces
  AttendanceSummary: RequestHandler = async (req, res) => {
    try {
      const year = req.params.year;
      const month = req.params.month || null; // Mes opcional

      // Llamar al servicio con el año y el mes
      const data = await this.categoryService.getAttendanceSummaryByYear(
        String(year),
        month ? month : null
      );

      // Enviar la respuesta
      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getAttendanceSummary:", error);

      // Manejar errores enviando la respuesta de error
      res.status(500).json({
        success: false,
        message: "Failed to fetch attendance summary",
      });
    }
  };

  // Obtener los datos de faltas por enfermedad y por vacaciones
  getAbsencesByType: RequestHandler = async (req, res) => {
    try {
      const year = req.params.year;
      const month = req.params.month || null; // Mes opcional

      // Llamar al servicio con el año y el mes
      const data = await this.categoryService.getAbsencesByTypeByYear(
        String(year),
        month ? month : null
      );

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getAbsencesByType:", error);

      res.status(500).json({
        success: false,
        message: "Failed to fetch absences by type",
      });
    }
  };

  // Obtener todos los empleados
  getAllEmployees: RequestHandler = async (req, res) => {
    try {
      const data = await this.categoryService.getAllEmployees();

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getAllEmployees:", error);

      res.status(500).json({
        success: false,
        message: "Failed to fetch all employees",
      });
    }
  };

  // Obtener empleado por id
  getIDEmployees: RequestHandler = async (req, res) => {
    try {
      const id = req.params.id;
      const data = await this.categoryService.getEmployeeById(id);

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getIDEmployees:", error);

      res.status(500).json({
        success: false,
        message: "Failed to fetch employee",
      });
    }
  };

  // Editar empleado por id
  updateEmployee: RequestHandler = async (req, res) => {
    try {
      const id = req.params.id;
      const formData = req.body;
      const data = await this.categoryService.updateEmployeeById(id, formData);

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in updateEmployee:", error);

      res.status(500).json({
        success: false,
        message: "Failed to update employee",
      });
    }
  };

  // Obtener todos los departamentos
  getAllDepartaments: RequestHandler = async (req, res) => {
    try {
      const data = await this.categoryService.getAllDepartaments();

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getAllDepartaments:", error);

      res.status(500).json({
        success: false,
        message: "Failed to fetch all departaments",
      });
    }
  };

  // Obtener empleados con vacaciones
  getAllVacations: RequestHandler = async (req, res) => {
    try {
      const data = await this.categoryService.getAllVacations();

      res.status(200).json({
        success: true,
        data,
      });
    } catch (error) {
      console.error("Error in getAllVacations:", error);

      res.status(500).json({
        success: false,
        message: "Failed to fetch all vacations",
      });
    }
  };
}
