// AttendanceRoutes.ts
import { Router } from "express";
import { AttendanceService } from "../../services/attendance.service";
import { AttendanceController } from "./controller";

export class AttendanceRoutes {
  static get routes(): Router {
    const router = Router();

    const attendanceService = new AttendanceService();
    const attendanceController = new AttendanceController(attendanceService);

    router.post(
      "/monthly-attendance-report",
      attendanceController.getMonthlyAttendanceReport
    );

    router.post(
      "/searchAttendance",
      attendanceController.searchAttendance
    );

    router.get(
      "/download-excel",
      attendanceController.downloadMonthlyAttendanceReport
    );


    router.get(
      "/total-dashboard",
      attendanceController.totalDashboardReport
    );

    router.get(
      "/attendance-summary/:year/:month?",
      attendanceController.AttendanceSummary
    );

    router.get(
      "/absences-by-type/:year/:month?",
      attendanceController.getAbsencesByType
    );

    return router;
  }
}
