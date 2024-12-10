// src/routes/exception.routes.ts

import express from "express";
import { ExceptionService } from "../../services/exception.service";
import { ExceptionController } from "./controller";

export class ExceptionRoutes {
  static get routes(): express.Router {
    const router = express.Router();
    const exceptionService = new ExceptionService();
    const exceptionController = new ExceptionController(exceptionService);

    router.get("/exception-report", exceptionController.getExceptionReport);

    return router;
  }
}
