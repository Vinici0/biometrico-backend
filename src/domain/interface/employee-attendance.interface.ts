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
  Entrada: string | null;
  Salida: string | null;
  Departamento?: string;
  TotalHorasRedondeadas: number;
  DiaSemana: 'Domingo' | 'Lunes' | 'Martes' | 'Miércoles' | 'Jueves' | 'Viernes' | 'Sábado';
  B?: 'B' | 'Z' | null;
  Vacaciones: string | null;
}

export interface AsistenciaData {
  [key: string]: EmployeeAttendance[];
}
