import { Router } from "express";
import { AttendanceRoutes } from "./reports/routes";
import { EmployeeRoutes } from "./employee/routes";
import { ExceptionRoutes } from "./exception/routes";
import { SettingRoutes } from "./setting/routes";

export class AppRoutes {
  static get routes(): Router {
    const router = Router();

    router.use("/api/employees", EmployeeRoutes.routes);
    router.use("/api/exceptions", ExceptionRoutes.routes);
    router.use("/api/reports", AttendanceRoutes.routes);
    router.use("/api/settings", SettingRoutes.routes);

    return router;
  }
}
