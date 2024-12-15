import { Request, RequestHandler, Response } from "express";
import { CustomError } from "../../domain/response/custom.error";
import { HttpResponseHandler } from "../../domain/response/http-response-handler";
import { SettingService } from "../../services/setting.service";

export class SettingController {
  constructor(public readonly categoryService: SettingService) {}

  getSettings: RequestHandler = (req: Request, res: Response) => {
    this.categoryService
      .getSettings()
      .then((data) => {
        return HttpResponseHandler.success(res, data);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  updateSettings: RequestHandler = (req: Request, res: Response) => {
    const { settings } = req.body;
    
    this.categoryService
      .updateSettings(settings)
      .then((data) => {
        return HttpResponseHandler.success(res, data);
      })
      .catch((error) => {
        return this.handleError(error, res);
      });
  };

  private handleError = (error: unknown, res: Response) => {
    if (error instanceof CustomError) {
      return HttpResponseHandler.error(
        res,
        error,
        error.message,
        error.statusCode
      );
    }

    console.log(`${error}`);
    return HttpResponseHandler.error(res, error);
  };
}
