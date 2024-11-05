import Sequelize from "sequelize";
import sequelize from "../config/database";
import { createExcelReport } from "../utils/report-xml";
import { EmployeeAttendance } from "../domain/interface/employee-attendance.interface";


interface AttendanceSearchParams {
  name: string;
  startDate: string;
  endDate: string;
  page: number;
  pageSize: number;
}
export class AttendanceService {
  constructor() {}

  public async getMonthlyAttendanceReport(
    startDate = "",
    endDate = "",
    page = 1,
    pageSize = 10
  ) {
    try {
      console.log( startDate,
        endDate, 
        page ,
        pageSize);
      
      const offset = (page - 1) * pageSize;

      const results = (await sequelize.query(
        ` 
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
              `,
        {
          replacements: { startDate, endDate, pageSize, offset },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as EmployeeAttendance[];

      return results;
    } catch (error) {
      console.error("Error fetching attendance report:", error);
      throw error;
    }
  }

  public async searchAttendanceByName(params: AttendanceSearchParams) {
    try {
      const { name = "", startDate = "", endDate = "", page = 1, pageSize = 10 } = params;
      const offset = (page - 1) * pageSize;
  
      const results = (await sequelize.query(
        ` 
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
              `,
        {
          replacements: { name: `%${name}%`, startDate, endDate, pageSize, offset },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as EmployeeAttendance[];
  
      return results;
    } catch (error) {
      console.error("Error fetching attendance report by name:", error);
      throw error;
    }
  }

  //getMonthlyAttendanceReport
  public async downloadMonthlyAttendanceReport({
    startDate = "2024-09-01",
    endDate = "2024-09-30",
  }) {
    try {
      console.log("startDate", startDate);
      console.log("endDate", endDate);

      const results = (await sequelize.query(
        `
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
                `,
        {
          replacements: { startDate, endDate },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as EmployeeAttendance[];

      const employeeResults: any = {};

      results.forEach((result) => {
        if (!employeeResults[result.id]) {
          employeeResults[result.id] = [];
        }
        employeeResults[result.id].push(result);
      });

      const buffer = await createExcelReport(9, 2024, employeeResults);
      return buffer;
    } catch (error) {
      console.error("Error fetching attendance report:", error);
      throw error;
    }
  }
}
