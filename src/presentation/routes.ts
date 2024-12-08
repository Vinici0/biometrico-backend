import { Router } from "express";
import { AttendanceRoutes } from "./reports/routes";
import { EmployeeRoutes } from "./employee/routes";


export class AppRoutes {
  static get routes(): Router {
    const router = Router();

    router.use("/api/reports", AttendanceRoutes.routes);
    router.use("/api/employees", EmployeeRoutes.routes);
    
    return router;
  }
}
