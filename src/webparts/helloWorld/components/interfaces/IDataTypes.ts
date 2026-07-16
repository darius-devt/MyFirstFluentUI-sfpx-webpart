/**
 * Define los tipos de datos que se usan en la tabla
 * Centraliza constantes y tipos reutilizables
 */
export interface IDataItem {
  key: string;
  id: number;
  nombre: string;
  departamento: string;
  email: string;
  estado: TEstado;
  fechaRegistro: string;
}

export const ESTADOS_DISPONIBLES = {
  ACTIVO: 'Activo',
  INACTIVO: 'Inactivo',
  PENDIENTE: 'Pendiente',
} as const;

export const DEPARTAMENTOS_DISPONIBLES = {
  VENTAS: 'Ventas',
  IT: 'IT',
  RECURSOS_HUMANOS: 'Recursos Humanos',
} as const;

export type TEstado = typeof ESTADOS_DISPONIBLES[keyof typeof ESTADOS_DISPONIBLES];
export type TDepartamento = typeof DEPARTAMENTOS_DISPONIBLES[keyof typeof DEPARTAMENTOS_DISPONIBLES];