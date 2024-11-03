import { Router } from "express";
import { AttendanceRoutes } from "./reports/routes";

export class AppRoutes {
  static get routes(): Router {
    const router = Router();

    router.use("/api/reports", AttendanceRoutes.routes);

    return router;
  }
}
