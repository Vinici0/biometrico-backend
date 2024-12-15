export interface ExceptionRecord {
    emp_pin: number;
    emp_firstname: string;
    emp_lastname: string;
    exception_date: string;
    starttime: string;
    endtime: string;
    total_horas_excepcion: number;
    tipo_de_excepcion: string;
  }
  
  export interface Attendance {
    id: number;
    Numero: string;
    Nombre: string;
    Fecha: string;
    Entrada: string;
    Salida: null | string;
    Departamento: string;
    TotalHorasRedondeadas: number | null;
    DiaSemana: string;
    status: string;
  }
  