import { Response } from "express";

export class HttpResponseHandler {
    public static success(res: Response, data: any = {}, message: string = "Success", statusCode: number = 200) {
      return res.status(statusCode).json({
        status: "success",
        message,
        data,
      });
    }
  
    public static created(res: Response, data: any = {}, message: string = "Created", statusCode: number = 201) {
      return res.status(statusCode).json({
        status: "success",
        message,
        data,
      });
    }
  
    public static error(res: Response, error: any = {}, message: string = "Internal Server Error", statusCode: number = 500) {
      return res.status(statusCode).json({
        status: "error",
        message,
        error,
      });
    }
  
    public static badRequest(res: Response, error: any = {}, message: string = "Bad Request", statusCode: number = 400) {
      return res.status(statusCode).json({
        status: "error",
        message,
        error,
      });
    }
  
    public static notFound(res: Response, message: string = "Resource Not Found", statusCode: number = 404) {
      return res.status(statusCode).json({
        status: "error",
        message,
      });
    }
  
    public static unauthorized(res: Response, message: string = "Unauthorized", statusCode: number = 401) {
      return res.status(statusCode).json({
        status: "error",
        message,
      });
    }
  }
  