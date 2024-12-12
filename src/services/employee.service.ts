// src/services/employee.service.ts

import Sequelize from "sequelize";
import sequelize from "../config/database";
import { Employee } from "../domain/interface/employee.interface";

interface EmployeeSearchParams {
  searchTerm?: string;
  page?: number;
  pageSize?: number;
}

export class EmployeeService {
  constructor() {}

  /**
   * Retrieves a list of employees with optional search filtering and pagination.
   *
   * @param {EmployeeSearchParams} params - The search parameters.
   * @returns {Promise<Employee[]>} - A promise that resolves to an array of employees.
   */
  public async getEmployees({
    searchTerm = "",
    page = 1,
    pageSize = 10,
  }: EmployeeSearchParams) {
    try {
      const offset = (page - 1) * pageSize;

      let query = `
        SELECT 
          e.id,
          e.emp_firstname || ' ' || e.emp_lastname AS "Nombre",
          d.dept_name AS "Departamento",
          e.emp_role AS "Rol",
          e.emp_email AS "Correo"
        FROM 
          hr_employee e
          LEFT JOIN hr_department d ON e.emp_dept = d.id
      `;

      const replacements: Record<string, any> = { pageSize, offset };

      // Solo agregar la condición WHERE si searchTerm tiene valor
      if (searchTerm) {
        query += `
          WHERE 
            e.emp_firstname || ' ' || e.emp_lastname LIKE :searchTerm 
            OR d.dept_name LIKE :searchTerm 
            OR e.emp_role LIKE :searchTerm 
            OR e.emp_email LIKE :searchTerm AND e.emp_active = 1
        `;
        replacements.searchTerm = `%${searchTerm}%`;
      }

      query += `
        ORDER BY "Nombre" ASC
        LIMIT :pageSize OFFSET :offset
      `;

      const results = await sequelize.query(query, {
        replacements,
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;
    } catch (error) {
      console.error("Error fetching employees:", error);
      throw error;
    }
  }

  getEmployeesByTerm = async (searchTerm: string) => {
    try {
      const query = `
        SELECT 
          id,
          emp_firstname || ' ' || emp_lastname AS "name",
          emp_role AS "role",
          emp_email AS "email"
        FROM 
          hr_employee
        WHERE 
          emp_firstname || ' ' || emp_lastname LIKE :searchTerm 
          OR emp_role LIKE :searchTerm 
          OR emp_email LIKE :searchTerm AND emp_active = 1
      `;

      const replacements = { searchTerm: `%${searchTerm}%` };

      const results = await sequelize.query(query, {
        replacements,
        type: Sequelize.QueryTypes.SELECT,
      });

      return results;
    } catch (error) {
      console.error("Error fetching employees by term:", error);
      throw error;
    }
  };
}
