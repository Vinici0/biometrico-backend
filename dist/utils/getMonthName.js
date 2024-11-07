"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getMonthName = void 0;
const getMonthName = (month) => {
    const monthNames = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ];
    return monthNames[month - 1];
};
exports.getMonthName = getMonthName;
