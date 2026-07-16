/**
 * Archivo: IHelloWorldState.ts
 * Descripción: Define el estado interno del componente HelloWorld.
 * Responsabilidad única: Definir la interfaz del estado del componente.
 * Nota importante: El estado es mutable y cambia durante la vida útil del componente.
 * Se inicializa en el constructor y se actualiza con setState().
 */
import { IColumn, IGroup } from '@fluentui/react';
import { IDataItem } from './IDataTypes';

/**
 * Interfaz para el estado del componente HelloWorld
 * Define el estado interno del componente
 * Estos datos pueden cambiar durante la ejecución (setState)
 */
export interface IHelloWorldState {
  isModalOpen: boolean; // Indica si el modal de información está abierto (true) o cerrado (false). Se utiliza para mostrar u ocultar el modal en la interfaz de usuario.
  items: IDataItem[]; // Array de elementos que se mostrarán en la tabla. Cada elemento debe cumplir con la interfaz IDataItem, que define las propiedades clave, id, nombre, departamento, email, estado y fechaRegistro.
  groups: IGroup[]; // Array de grupos para la tabla. Cada grupo define un conjunto de elementos que comparten una propiedad común (por ejemplo, departamento). Esto permite que la tabla muestre los elementos agrupados visualmente.
  columns: IColumn[]; // Array de columnas que definen la estructura de la tabla. Cada columna debe cumplir con la interfaz IColumn de Fluent UI, que define propiedades como key, name, fieldName, minWidth, maxWidth, isResizable, isSorted y onRender. Esto permite personalizar la apariencia y el comportamiento de cada columna en la tabla.
}
