export interface DataResponse {
  status: string;
  message: string;
  data: EmployeeAttendance[];
}

export interface EmployeeAttendance {
  id: number;
  Numero: string;
  Nombre: string;
  Fecha: string;
  Entrada: string;
  Salida: string;
  Departamento: string;
  TotalHorasRedondeadas: number;
  DiaSemana: string;
}

export interface AsistenciaData {
  [key: string]: EmployeeAttendance[];
}
