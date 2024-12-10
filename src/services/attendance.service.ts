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
  constructor() { }

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
            ea.paycode_id,
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
                ELSE NULL
            END AS "TipoB",
            CASE
                WHEN COUNT(p.punch_time) = 0 THEN 'Z'
                ELSE NULL
            END AS "TipoZ",
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
            LEFT JOIN att_exceptionassign ea ON ea.employee_id = e.id

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



  // ==========================================================
  // Métodos adicionales para el total conteo para dashboard

  public async totalDashboardReport() {
    try {
      const queries = {
        employee: {
          total: `SELECT COUNT(*) AS total FROM hr_employee;`,
          active: `SELECT COUNT(*) AS total FROM hr_employee WHERE emp_active = 1;`,
          inactive: `SELECT COUNT(*) AS total FROM hr_employee WHERE emp_active = 0;`,
        },
        assistant: {
          total: `
            SELECT COUNT(DISTINCT ADDL.employee_id) AS employee_count
            FROM att_day_details AS ADDL
            INNER JOIN hr_employee AS HE ON ADDL.employee_id = HE.id
            WHERE DATE(ADDL.checkin) = CURRENT_DATE;
          `,
          support: `
            SELECT COUNT(DISTINCT ADDL.employee_id) AS employee_count
            FROM att_day_details AS ADDL
            INNER JOIN hr_employee AS HE ON ADDL.employee_id = HE.id
            WHERE HE.emp_dept = 1 AND DATE(ADDL.checkin) = CURRENT_DATE;
          `,
          normal: `
            SELECT COUNT(DISTINCT ADDL.employee_id) AS employee_count
            FROM att_day_details AS ADDL
            INNER JOIN hr_employee AS HE ON ADDL.employee_id = HE.id
            WHERE HE.emp_dept = 5 AND DATE(ADDL.checkin) = CURRENT_DATE;
          `,
        },
        vacation: {
          total: `
            SELECT COUNT(*) AS total
            FROM att_paycode AP
            INNER JOIN att_exceptionassign AEA ON AP.id = AEA.paycode_id
            INNER JOIN hr_employee HE ON AEA.employee_id = HE.id
            WHERE DATE(AEA.exception_date) = CURRENT_DATE
            AND HE.emp_active = 1;
          `,
        },
      };

      const results = await Promise.all(
        Object.entries(queries).flatMap(([groupKey, groupQueries]) =>
          Object.entries(groupQueries).map(async ([key, query]) => {
            const result = await sequelize.query<{ total?: number; employee_count?: number }>(query, {
              type: Sequelize.QueryTypes.SELECT,
            });

            // Validar la existencia de las claves esperadas
            const value = result[0]?.total ?? result[0]?.employee_count;

            if (value !== undefined) {
              return { [`${groupKey}.${key}`]: value };
            } else {
              throw new Error(`Query for ${key} did not return a valid result.`);
            }
          })
        )
      );

      // Combina los resultados en un solo objeto
      const combinedResults = results.reduce((acc, curr) => ({ ...acc, ...curr }), {});

      // Agregar cálculo para earrings.attendance
      const assistantTotal = combinedResults["assistant.total"] || 0;
      const employeeActive = combinedResults["employee.active"] || 0;
      combinedResults["earrings.attendance"] = Math.abs(assistantTotal - employeeActive);

      return combinedResults;
    } catch (error) {
      console.error("Error fetching total dashboard report:", error);
      throw error;
    }
  }


  // Total de asistencias, ausencias y atrasos
  public async getAttendanceSummaryByYear(year: string, month: string | null) {
    try {
      const query = `
      SELECT
          HE.emp_dept AS department,
          strftime('%Y-%m', ADDL.att_date) AS month,
          COUNT(DISTINCT CASE WHEN ADDL.checkin IS NOT NULL THEN ADDL.employee_id || ADDL.att_date END) AS total_attendances,
          COUNT(DISTINCT CASE WHEN ADDL.checkin IS NULL THEN ADDL.employee_id || ADDL.att_date END) AS total_absences,
          COUNT(DISTINCT CASE
                            WHEN ADDL.checkin IS NOT NULL AND ADDL.checkin > AST.shift_start THEN ADDL.employee_id || ADDL.att_date
          END) AS total_late
      FROM
          att_day_details AS ADDL
          LEFT JOIN hr_employee AS HE ON ADDL.employee_id = HE.id
          LEFT JOIN att_shift AS AST ON AST.id = ADDL.shift_ID
      WHERE 
          strftime('%Y', ADDL.att_date) = :year
          ${month ? `AND strftime('%m', ADDL.att_date) = :month` : ""}
      GROUP BY
          HE.emp_dept,
          strftime('%Y-%m', ADDL.att_date)
      ORDER BY
          HE.emp_dept,
          strftime('%Y-%m', ADDL.att_date);
      `;

      const replacements: { year: string; month?: string } = { year };
      if (month) {
        replacements.month = month.padStart(2, "0");
      }

      const results = await sequelize.query(query, {
        replacements,
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;
    } catch (error) {
      console.error("Error fetching attendance summary by year and department:", error);
      throw error;
    }
  }


  // Obtener los datos de faltas por enfermedad y por vacaciones
  public async getAbsencesByTypeByYear(year: string, month: string | null) {
    try {
      const query = `
      SELECT 
          HE.emp_dept AS department,
          strftime('%Y-%m', ADS.att_date) AS month,
          COUNT(CASE WHEN APC.id = 12 THEN ADS.id END) AS sickness_absences,
          COUNT(CASE WHEN APC.id = 11 THEN ADS.id END) AS vacation_absences
      FROM 
          att_day_summary AS ADS
      LEFT JOIN att_paycode AS APC 
          ON ADS.paycode_id = APC.id
      LEFT JOIN hr_employee AS HE 
          ON ADS.employee_id = HE.id
      WHERE 
          strftime('%Y', ADS.att_date) = :year
          ${month ? `AND strftime('%m', ADS.att_date) = :month` : ""}
      GROUP BY 
          HE.emp_dept,
          strftime('%Y-%m', ADS.att_date)
      ORDER BY 
          HE.emp_dept,
          strftime('%Y-%m', ADS.att_date);
    `;

      const replacements: { year: string; month?: string } = { year }; // Define reemplazos con tipo opcional
      if (month) {
        replacements.month = month.padStart(2, "0"); // Asegura formato MM
      }

      const results = await sequelize.query(query, {
        replacements,
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;
    } catch (error) {
      console.error("Error fetching absences by type for year:", error);
      throw error;
    }
  }

  // Obtener todos los mepleados
  getAllEmployees() {
    try {

      const query = `
        SELECT HE.id, HE.emp_firstname, HE.emp_lastname, He.emp_email, HE.emp_dept, HD.id AS id_dept, HD.dept_name, HE.emp_active FROM hr_employee AS HE
        INNER JOIN hr_department AS HD ON HE.emp_dept = HD.id
      `;

      const results = sequelize.query(query, {
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;

    } catch (error) {
      console.error("Error fetching all employees:", error);
      throw error;
    }

  }

  // Obtener empleado por id
  getEmployeeById(id: string) {
    try {

      const query = `
        SELECT HE.emp_firstname, HE.emp_lastname, HE.emp_email, HE.emp_dept, HD.dept_name, HE.emp_active FROM hr_employee AS HE
        INNER JOIN hr_department AS HD ON HE.emp_dept = HD.id
        WHERE HE.id = :id;
      `;

      const results = sequelize.query(query, {
        replacements: { id },
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;

    } catch (error) {
      console.error("Error fetching employee by id:", error);
      throw error;
    }
  }

  // Editar empleado por id
  updateEmployeeById(id: string, formData: any) {
    try {

      const {
        emp_firstname,
        emp_lastname,
        emp_email,
        id_dept,
        emp_active,
      } = formData;

      const query = `
        UPDATE hr_employee
          SET 
            emp_firstname = :emp_firstname,
            emp_lastname = :emp_lastname,
            emp_email = :emp_email,
            emp_dept = :id_dept,
            emp_active = :emp_active
        WHERE 
          id = :id;
      `;

      const results
        = sequelize.query(query, {
          replacements: {
            emp_firstname,
            emp_lastname,
            emp_email,
            id_dept,
            emp_active,
            id
          },
          type: Sequelize.QueryTypes.UPDATE,
        });

      return results;

    } catch (error) {

      console.error("Error updating employee:", error);
      throw error;
    }
  }

  // Obtener todos los departamentos
  getAllDepartaments() {
    try {

      const query = `
        SELECT * FROM hr_department;
      `;

      const results = sequelize.query(query, {
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;

    } catch (error) {
      console.error("Error fetching all departaments:", error);
      throw error;
    }
  }

  // Obtener todos los registros de vacaciones
  getAllVacations() {
    try {
      const query = `
          SELECT
              HE.emp_pin AS emp_pin,
              HE.emp_firstname AS emp_firstname,
              HE.emp_lastname AS emp_lastname,
              AEA.exception_date AS exception_date,
              AEA.starttime AS starttime,
              AEA.endtime AS endtime,
              ROUND((julianday(AEA.endtime) - julianday(AEA.starttime)) * 24, 2) AS total_horas_excepcion,
              CASE
                  WHEN AP.pc_desc LIKE '%vacation%' THEN 'Vacaciones'
                  WHEN AP.pc_desc LIKE '%sick%' THEN 'Enfermedad'
                  ELSE AP.pc_desc
                  END AS tipo_de_excepcion
          FROM
              att_paycode AP
                  INNER JOIN
              att_exceptionassign AEA ON AP.id = AEA.paycode_id
                  INNER JOIN
              hr_employee HE ON AEA.employee_id = HE.id
          WHERE
            DATE(AEA.exception_date) = CURRENT_DATE
            AND HE.emp_active = 1
          ORDER BY
              HE.emp_pin;
      `;

      const results = sequelize.query(query, {
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;

    } catch (error) {
      console.error("Error fetching all vacations:", error);
      throw error;
    }
  }

}
