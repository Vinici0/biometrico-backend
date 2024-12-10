export interface DataResponse {
  status: string;
  message: string;
  data: EmployeeAttendance[];
}

export interface EmployeeAttendance {
  EmpleadoID: number; // ID del empleado
  Nombre: string; // Nombre completo del empleado
  Numero: string; // Número de empleado
  PaycodeID: number | null; // ID del código de pago (puede ser null)
  Fecha: string; // Fecha (formato ISO: YYYY-MM-DD)
  Entrada: string | null; // Hora de entrada (formato HH:mm) o null si no hay registros
  Salida: string | null; // Hora de salida (formato HH:mm) o null si no hay registros
  Departamento: string | null; // Nombre del departamento (puede ser null)
  TotalHorasRedondeadas: number | null; // Total de horas trabajadas redondeadas, o null
  DiaSemana: string; // Nombre del día de la semana
  TipoB: string | null; // 'B' si hay un solo registro de hora, null de lo contrario
  TipoZ: string | null; // 'Z' si no hay registros de hora, null de lo contrario
  Vacaciones: string; // 'Si' si el paycode_id es 12, 'No' de lo contrario
  Enfermedad: string; // 'Si' si el paycode_id es 11, 'No' de lo contrario
}

export interface AsistenciaData {
  [key: string]: EmployeeAttendance[];
}
