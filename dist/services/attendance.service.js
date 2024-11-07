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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AttendanceService = void 0;
const sequelize_1 = __importDefault(require("sequelize"));
const database_1 = __importDefault(require("../config/database"));
const report_xml_1 = require("../utils/report-xml");
class AttendanceService {
    constructor() { }
    getMonthlyAttendanceReport() {
        return __awaiter(this, arguments, void 0, function* (startDate = "", endDate = "", page = 1, pageSize = 10) {
            try {
                console.log(startDate, endDate, page, pageSize);
                const offset = (page - 1) * pageSize;
                const results = (yield database_1.default.query(` 
        SELECT e.id, e.emp_code                               AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        date(p.punch_time)                       AS "Fecha",
        MIN(time(p.punch_time))                  AS "Entrada",
        MAX(time(p.punch_time))                  AS "Salida",
        dp.dept_name                             AS "Departamento",
        ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) *
              24)                                AS "TotalHorasRedondeadas",
        CASE strftime('%w', date(p.punch_time))
            WHEN '0' THEN 'Domingo'
            WHEN '1' THEN 'Lunes'
            WHEN '2' THEN 'Martes'
            WHEN '3' THEN 'Miércoles'
            WHEN '4' THEN 'Jueves'
            WHEN '5' THEN 'Viernes'
            WHEN '6' THEN 'Sábado'
            END                                  AS "DiaSemana"
  
          FROM att_punches p
                  JOIN
              hr_employee e ON p.emp_id = e.id
                  LEFT JOIN
              hr_department dp ON e.emp_dept = dp.id
          WHERE date(p.punch_time) BETWEEN :startDate AND :endDate
          GROUP BY date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
          ORDER BY date(p.punch_time)
          LIMIT :pageSize OFFSET :offset;
              `, {
                    replacements: { startDate, endDate, pageSize, offset },
                    type: sequelize_1.default.QueryTypes.SELECT,
                }));
                return results;
            }
            catch (error) {
                console.error("Error fetching attendance report:", error);
                throw error;
            }
        });
    }
    searchAttendanceByName(params) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const { name = "", startDate = "", endDate = "", page = 1, pageSize = 10 } = params;
                const offset = (page - 1) * pageSize;
                const results = (yield database_1.default.query(` 
        SELECT e.id, e.emp_code                               AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        date(p.punch_time)                       AS "Fecha",
        MIN(time(p.punch_time))                  AS "Entrada",
        MAX(time(p.punch_time))                  AS "Salida",
        dp.dept_name                             AS "Departamento",
        ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) *
              24)                                AS "TotalHorasRedondeadas",
        CASE strftime('%w', date(p.punch_time))
            WHEN '0' THEN 'Domingo'
            WHEN '1' THEN 'Lunes'
            WHEN '2' THEN 'Martes'
            WHEN '3' THEN 'Miércoles'
            WHEN '4' THEN 'Jueves'
            WHEN '5' THEN 'Viernes'
            WHEN '6' THEN 'Sábado'
            END                                  AS "DiaSemana"
  
          FROM att_punches p
                  JOIN
              hr_employee e ON p.emp_id = e.id
                  LEFT JOIN
              hr_department dp ON e.emp_dept = dp.id
          WHERE (e.emp_firstname || ' ' || e.emp_lastname LIKE :name)
            AND (date(p.punch_time) BETWEEN :startDate AND :endDate)
          GROUP BY date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
          ORDER BY date(p.punch_time)
          LIMIT :pageSize OFFSET :offset;
              `, {
                    replacements: { name: `%${name}%`, startDate, endDate, pageSize, offset },
                    type: sequelize_1.default.QueryTypes.SELECT,
                }));
                return results;
            }
            catch (error) {
                console.error("Error fetching attendance report by name:", error);
                throw error;
            }
        });
    }
    //getMonthlyAttendanceReport
    downloadMonthlyAttendanceReport(_a) {
        return __awaiter(this, arguments, void 0, function* ({ startDate = "2024-09-01", endDate = "2024-09-30", }) {
            try {
                console.log("startDate", startDate);
                console.log("endDate", endDate);
                const results = (yield database_1.default.query(`
                  SELECT e.id, e.emp_code                               AS "Numero",
        e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
        date(p.punch_time)                       AS "Fecha",
        MIN(time(p.punch_time))                  AS "Entrada",
        MAX(time(p.punch_time))                  AS "Salida",
        dp.dept_name                             AS "Departamento",
        ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) *
              24)                                AS "TotalHorasRedondeadas",
        CASE strftime('%w', date(p.punch_time))
            WHEN '0' THEN 'Domingo'
            WHEN '1' THEN 'Lunes'
            WHEN '2' THEN 'Martes'
            WHEN '3' THEN 'Miércoles'
            WHEN '4' THEN 'Jueves'
            WHEN '5' THEN 'Viernes'
            WHEN '6' THEN 'Sábado'
            END                                  AS "DiaSemana"

          FROM att_punches p
                  JOIN
              hr_employee e ON p.emp_id = e.id
                  LEFT JOIN
              hr_department dp ON e.emp_dept = dp.id
          WHERE date(p.punch_time) BETWEEN :startDate AND :endDate
          GROUP BY date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
          ORDER BY date(p.punch_time);
                `, {
                    replacements: { startDate, endDate },
                    type: sequelize_1.default.QueryTypes.SELECT,
                }));
                const employeeResults = {};
                results.forEach((result) => {
                    if (!employeeResults[result.id]) {
                        employeeResults[result.id] = [];
                    }
                    employeeResults[result.id].push(result);
                });
                const buffer = yield (0, report_xml_1.createExcelReport)(9, 2024, employeeResults);
                return buffer;
            }
            catch (error) {
                console.error("Error fetching attendance report:", error);
                throw error;
            }
        });
    }
}
exports.AttendanceService = AttendanceService;
