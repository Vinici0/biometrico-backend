"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.AttendanceRoutes = void 0;
// AttendanceRoutes.ts
const express_1 = require("express");
const attendance_service_1 = require("../../services/attendance.service");
const controller_1 = require("./controller");
class AttendanceRoutes {
    static get routes() {
        const router = (0, express_1.Router)();
        const attendanceService = new attendance_service_1.AttendanceService();
        const attendanceController = new controller_1.AttendanceController(attendanceService);
        router.post("/monthly-attendance-report", attendanceController.getMonthlyAttendanceReport);
        router.post("/searchAttendance", attendanceController.searchAttendance);
        router.get("/download-excel", attendanceController.downloadMonthlyAttendanceReport);
        return router;
    }
}
exports.AttendanceRoutes = AttendanceRoutes;
