import Sequelize from "sequelize";
import sequelize from "../config/database";
import { createExcelReport } from "../utils/report-xml";
import { AsistenciaData, EmployeeAttendance } from "../domain/interface/employee-attendance.interface";

interface AttendanceSearchParams {
  name: string;
  startDate: string;
  endDate: string;
  page: number;
  pageSize: number;
  department: string;
  status: string; // Nuevo parámetro
  employeeId: string;
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
      console.log(startDate, endDate, page, pageSize);

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
          WHERE date(p.punch_time) BETWEEN :startDate AND :endDate AND dp.dept_name  = 'Personal.HeH'
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

  public async searchAttendance(params: AttendanceSearchParams) {
    try {
      const {
        name = "",
        startDate = "",
        endDate = "",
        page = 1,
        pageSize = 10,
        department = "",
        status = "Todos",
        employeeId = "",
      } = params;

      // Validar employeeId si se proporciona
      if (employeeId && isNaN(Number(employeeId))) {
        throw new Error("Invalid employeeId provided.");
      }

      const offset = (page - 1) * pageSize;

      // Base query
      let baseQuery = `
          FROM att_punches p
          JOIN hr_employee e ON p.emp_id = e.id
          LEFT JOIN hr_department dp ON e.emp_dept = dp.id
        `;

      // Construcción dinámica del WHERE y HAVING
      const whereClauses: string[] = [];
      const havingClauses: string[] = [];
      const replacements: Record<string, any> = {
        pageSize,
        offset,
      };

      if (name) {
        whereClauses.push(
          `(e.emp_firstname || ' ' || e.emp_lastname LIKE :name)`
        );
        replacements.name = `%${name}%`;
      }

      if (startDate && endDate) {
        whereClauses.push(
          `(date(p.punch_time) BETWEEN :startDate AND :endDate)`
        );
        replacements.startDate = startDate;
        replacements.endDate = endDate;
      }

      if (department) {
        whereClauses.push(`dp.dept_name = :department`);
        replacements.department = department;
      }

      if (employeeId) {
        // Agregar filtro por employeeId
        whereClauses.push(`e.id = :employeeId`);
        replacements.employeeId = Number(employeeId);
      }

      // Construir la cláusula WHERE
      if (whereClauses.length > 0) {
        baseQuery += ` WHERE ${whereClauses.join(" AND ")}`;
      }

      // Filtrar por status usando HAVING
      if (status === "Completo") {
        // Asegurarse de que hay más de un registro (entrada y salida diferentes)
        havingClauses.push(`COUNT(p.punch_time) > 1`);
      } else if (status === "Pendiente") {
        // Solo hay un registro (solo entrada)
        havingClauses.push(`COUNT(p.punch_time) = 1`);
      }

      // Construir la cláusula HAVING
      let havingClause = "";
      if (havingClauses.length > 0) {
        havingClause = `HAVING ${havingClauses.join(" AND ")}`;
      }

      // Consulta para obtener los datos paginados
      const dataQuery = `
          SELECT 
            e.id, 
            e.emp_code AS "Numero",
            e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
            date(p.punch_time) AS "Fecha",
            MIN(time(p.punch_time)) AS "Entrada",
            CASE 
              WHEN COUNT(p.punch_time) > 1 THEN MAX(time(p.punch_time)) 
              ELSE NULL 
            END AS "Salida",
            dp.dept_name AS "Departamento",
            CASE 
              WHEN COUNT(p.punch_time) > 1 THEN ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24) 
              ELSE NULL 
            END AS "TotalHorasRedondeadas",
            CASE strftime('%w', date(p.punch_time))
                WHEN '0' THEN 'Domingo'
                WHEN '1' THEN 'Lunes'
                WHEN '2' THEN 'Martes'
                WHEN '3' THEN 'Miércoles'
                WHEN '4' THEN 'Jueves'
                WHEN '5' THEN 'Viernes'
                WHEN '6' THEN 'Sábado'
            END AS "DiaSemana",
            CASE 
              WHEN COUNT(p.punch_time) > 1 THEN 'Completo'
              ELSE 'Pendiente'
            END AS "status"
          ${baseQuery}
          GROUP BY date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
          ${havingClause}
          ORDER BY date(p.punch_time)
          LIMIT :pageSize OFFSET :offset
        `;
      console.log(dataQuery);

      // Consulta para obtener el total de registros considerando los grupos y el HAVING
      const countQuery = `
          SELECT COUNT(*) as total
          FROM (
            SELECT 1
            ${baseQuery}
            GROUP BY date(p.punch_time), e.emp_code, e.emp_firstname, e.emp_lastname, dp.dept_name
            ${havingClause}
          ) as subquery
        `;
      console.log(countQuery);

      // Ejecutar ambas consultas en paralelo
      const [results, countResults] = await Promise.all([
        sequelize.query(dataQuery, {
          replacements,
          type: Sequelize.QueryTypes.SELECT,
        }) as Promise<EmployeeAttendance[]>,
        sequelize.query(countQuery, {
          replacements,
          type: Sequelize.QueryTypes.SELECT,
        }) as Promise<{ total: number }[]>,
      ]);

      // Obtener el total de registros
      const total = countResults[0].total;

      return { attendances: results, total };
    } catch (error) {
      console.error("Error fetching attendance report:", error);
      throw error;
    }
  }

  //getMonthlyAttendanceReport
  public async downloadMonthlyAttendanceReport({
    startDate = "2024-09-01",
    endDate = "2024-09-30",
  }) {
    try {
      const query = `
        WITH RECURSIVE dates AS (
            SELECT DATE(:startDate) AS att_date
            UNION ALL
            SELECT DATE(att_date, '+1 day')
            FROM dates
            WHERE att_date < DATE(:endDate)
        )
        SELECT
            e.id,
            e.emp_code AS "Numero",
            e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
            d.att_date AS "Fecha",
            MIN(strftime('%H:%M', p.punch_time)) AS "Entrada",
            MAX(strftime('%H:%M', p.punch_time)) AS "Salida",
            dp.dept_name AS "Departamento",
            ROUND((julianday(MAX(p.punch_time)) - julianday(MIN(p.punch_time))) * 24) AS "TotalHorasRedondeadas",
            CASE
                strftime('%w', d.att_date)
                    WHEN '0' THEN 'Domingo'
                    WHEN '1' THEN 'Lunes'
                    WHEN '2' THEN 'Martes'
                    WHEN '3' THEN 'Miércoles'
                    WHEN '4' THEN 'Jueves'
                    WHEN '5' THEN 'Viernes'
                    WHEN '6' THEN 'Sábado'
            END AS "DiaSemana",
            CASE
                WHEN COUNT(p.punch_time) = 1 THEN 'B'
                WHEN COUNT(p.punch_time) = 0 THEN 'Z'
                ELSE NULL
            END AS "B",
            CASE
                WHEN pcd.pc_desc LIKE 'vacat%' THEN 'Si'
                ELSE 'No'
            END AS "Vacaciones"
        FROM
            hr_employee e
            CROSS JOIN dates d
            LEFT JOIN att_punches p ON e.id = p.emp_id AND DATE(p.punch_time) = d.att_date
            LEFT JOIN hr_department dp ON e.emp_dept = dp.id
            LEFT JOIN att_day_summary ds ON ds.employee_id = e.id AND d.att_date = ds.att_date
            LEFT JOIN att_paycode pcd ON pcd.id = ds.paycode_id
        GROUP BY
            e.id,
            d.att_date,
            e.emp_code,
            e.emp_firstname,
            e.emp_lastname,
            dp.dept_name,
            pcd.pc_desc
        ORDER BY
            d.att_date,
            e.emp_lastname,
            e.emp_firstname; 
      `;
  
      // Ejecutar la consulta con reemplazos para startDate y endDate
      const results = (await sequelize.query(
        query,
        {
          replacements: { startDate, endDate },
          type: Sequelize.QueryTypes.SELECT,
        }
      )) as EmployeeAttendance[];
  
      const employeeResults: AsistenciaData = {};
  
      results.forEach((result) => {
        if (!employeeResults[result.id]) {
          employeeResults[result.id] = [];
        }
        employeeResults[result.id].push(result);
      });
  
      const buffer = await createExcelReport(
        startDate,
        endDate,
        employeeResults
      );
      return buffer;
    } catch (error) {
      console.error("Error fetching attendance report:", error);
      throw error;
    }
  }
}
