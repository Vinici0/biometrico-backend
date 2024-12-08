"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AttendanceController = void 0;
const http_response_handler_1 = require("../../domain/response/http-response-handler");
const custom_error_1 = require("../../domain/response/custom.error");
class AttendanceController {
    constructor(categoryService) {
        this.categoryService = categoryService;
        this.getMonthlyAttendanceReport = (req, res) => {
            const { startDate, endDate, page, pageSize } = req.body;
            this.categoryService
                .getMonthlyAttendanceReport(startDate, endDate, page, pageSize)
                .then((data) => {
                return http_response_handler_1.HttpResponseHandler.success(res, data);
            })
                .catch((error) => {
                return this.handleError(error, res);
            });
        };
        this.searchAttendance = (req, res) => {
            const { name, page, pageSize, startDate, endDate, department } = req.body;
            this.categoryService
                .searchAttendance({ name, page, pageSize, startDate, endDate, department })
                .then((data) => {
                return http_response_handler_1.HttpResponseHandler.success(res, data);
            })
                .catch((error) => {
                return this.handleError(error, res);
            });
        };
        //descargar reporte de asistencia xml
        this.downloadMonthlyAttendanceReport = (req, res) => __awaiter(this, void 0, void 0, function* () {
            const { startDate, endDate } = req.query;
            try {
                const buffer = yield this.categoryService.downloadMonthlyAttendanceReport({
                    startDate: startDate,
                    endDate: endDate,
                });
                // Configura los encabezados de la respuesta para forzar la descarga
                res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                res.setHeader("Content-Disposition", `attachment; filename="ReporteAsistencia_Firma.xlsx"`);
                // Envía el archivo como un adjunto en la respuesta
                res.send(buffer);
            }
            catch (error) {
                this.handleError(error, res);
            }
        });
        this.handleError = (error, res) => {
            if (error instanceof custom_error_1.CustomError) {
                return http_response_handler_1.HttpResponseHandler.error(res, error, error.message, error.statusCode);
            }
            console.log(`${error}`);
            return http_response_handler_1.HttpResponseHandler.error(res, error);
        };
    }
}
exports.AttendanceController = AttendanceController;
