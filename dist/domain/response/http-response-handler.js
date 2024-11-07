"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.HttpResponseHandler = void 0;
class HttpResponseHandler {
    static success(res, data = {}, message = "Success", statusCode = 200) {
        return res.status(statusCode).json({
            status: "success",
            message,
            data,
        });
    }
    static created(res, data = {}, message = "Created", statusCode = 201) {
        return res.status(statusCode).json({
            status: "success",
            message,
            data,
        });
    }
    static error(res, error = {}, message = "Internal Server Error", statusCode = 500) {
        return res.status(statusCode).json({
            status: "error",
            message,
            error,
        });
    }
    static badRequest(res, error = {}, message = "Bad Request", statusCode = 400) {
        return res.status(statusCode).json({
            status: "error",
            message,
            error,
        });
    }
    static notFound(res, message = "Resource Not Found", statusCode = 404) {
        return res.status(statusCode).json({
            status: "error",
            message,
        });
    }
    static unauthorized(res, message = "Unauthorized", statusCode = 401) {
        return res.status(statusCode).json({
            status: "error",
            message,
        });
    }
}
exports.HttpResponseHandler = HttpResponseHandler;
