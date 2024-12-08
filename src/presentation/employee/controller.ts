// src/controllers/employee.controller.ts

import { Request, Response, RequestHandler } from "express";
import { EmployeeService } from "../../services/employee.service";
import { HttpResponseHandler } from "../../domain/response/http-response-handler";
import { CustomError } from "../../domain/response/custom.error";

export class EmployeeController {
  constructor(private employeeService: EmployeeService) {}

  // Método para obtener empleados con paginación y búsqueda
  public getEmployees: RequestHandler = (req: Request, res: Response) => {
    const { searchTerm, page, pageSize } = req.query;
    console.log(searchTerm, page, pageSize);

    this.employeeService
      .getEmployees({
        searchTerm: searchTerm as string,
        page: Number(page),
        pageSize: Number(pageSize),
      })
      .then((employees) => {
        return HttpResponseHandler.success(res, employees);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  // Método para obtener empleados por término de búsqueda
  public getEmployeesByTerm: RequestHandler = (req: Request, res: Response) => {
    const { searchTerm } = req.query;
    console.log(searchTerm);

    this.employeeService
      .getEmployeesByTerm(searchTerm as string)
      .then((employees) => {
        return HttpResponseHandler.success(res, employees);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  // Manejo de errores
  private handleError = (error: unknown, res: Response) => {
    if (error instanceof CustomError) {
      return HttpResponseHandler.error(
        res,
        error,
        error.message,
        error.statusCode
      );
    }

    console.error("Error in EmployeeController:", error);
    return HttpResponseHandler.error(res, error);
  };
}
