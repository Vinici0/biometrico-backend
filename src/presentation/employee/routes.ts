// src/routes/employee.routes.ts

import { Router } from "express";
import { EmployeeController } from "./controller";
import { EmployeeService } from "../../services/employee.service";

export class EmployeeRoutes {
  static get routes(): Router {
    const router = Router();

    const employeeService = new EmployeeService();
    const employeeController = new EmployeeController(employeeService);

    router.get("/", employeeController.getEmployees);
    router.get("/search", employeeController.getEmployeesByTerm);

    return router;
  }
}
