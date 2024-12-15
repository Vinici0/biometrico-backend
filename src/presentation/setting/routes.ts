// AttendanceRoutes.ts
import { Router } from "express";
import { AttendanceService } from "../../services/attendance.service";
import { SettingService } from "../../services/setting.service";
import { SettingController } from "./controller";

export class SettingRoutes {
  static get routes(): Router {
    const router = Router();

    const settingService = new SettingService();
    const settingController = new SettingController(settingService);

    router.get("/", settingController.getSettings);
    router.put("/", settingController.updateSettings);

    return router;
  }
}
