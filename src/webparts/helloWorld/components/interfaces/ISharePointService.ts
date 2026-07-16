/**
 * Archivo: ISharePointService.ts
 * Descripción: Define la interfaz para el servicio de SharePoint.
 * 
 * Esta interfaz garantiza que cualquier implementación del servicio
 * cumpla con los métodos necesarios para interactuar con SharePoint.
 * 
 * Autor: Manus AI
 * Fecha: 2026
 */

import { IDataItem } from './IDataTypes';

/**
 * Interfaz para el servicio de SharePoint
 * Define los métodos que el servicio debe implementar
 */
export interface ISharePointService {
  /**
   * Obtiene todos los elementos de una lista de SharePoint
   * 
   * @param listName - Nombre o ID de la lista en SharePoint
   * @returns Promise con un array de IDataItem
   * 
   * Ejemplo de uso:
   * const items = await sharePointService.getListItems('Empleados');
   */
  getListItems(listName: string): Promise<IDataItem[]>;

  /**
   * Obtiene elementos de una lista con filtros aplicados
   * 
   * @param listName - Nombre o ID de la lista
   * @param filterQuery - Consulta CAML o OData para filtrar ($filter=...)
   * @returns Promise con un array de IDataItem filtrados
   * 
   * Ejemplo de uso:
   * const activeItems = await sharePointService.getListItemsFiltered(
   *   'Empleados',
   *   "Estado eq 'Activo'"
   * );
   */
  getListItemsFiltered(listName: string, filterQuery: string): Promise<IDataItem[]>;

  /**
   * Crea un nuevo elemento en una lista de SharePoint
   * 
   * @param listName - Nombre o ID de la lista
   * @param item - Objeto con los datos del nuevo elemento
   * @returns Promise con el elemento creado (incluyendo ID)
   * 
   * Ejemplo de uso:
   * const newEmployee = await sharePointService.createListItem(
   *   'Empleados',
   *   { nombre: 'Juan', departamento: 'IT', email: 'juan@empresa.com' }
   * );
   */
  createListItem(listName: string, item: Partial<IDataItem>): Promise<IDataItem>;

  /**
   * Actualiza un elemento existente en una lista
   * 
   * @param listName - Nombre o ID de la lista
   * @param itemId - ID del elemento a actualizar
   * @param item - Objeto con los datos a actualizar
   * @returns Promise con el elemento actualizado
   * 
   * Ejemplo de uso:
   * const updated = await sharePointService.updateListItem(
   *   'Empleados',
   *   1,
   *   { estado: 'Inactivo' }
   * );
   */
  updateListItem(listName: string, itemId: number, item: Partial<IDataItem>): Promise<IDataItem>;

  /**
   * Elimina un elemento de una lista
   * 
   * @param listName - Nombre o ID de la lista
   * @param itemId - ID del elemento a eliminar
   * @returns Promise<void> - Se resuelve cuando se completa la eliminación
   * 
   * Ejemplo de uso:
   * await sharePointService.deleteListItem('Empleados', 1);
   */
  deleteListItem(listName: string, itemId: number): Promise<void>;

  /**
   * Obtiene elementos con paginación
   * 
   * @param listName - Nombre o ID de la lista
   * @param pageNumber - Número de página (comenzando en 1)
   * @param pageSize - Cantidad de elementos por página
   * @returns Promise con la respuesta paginada
   * 
   * Ejemplo de uso:
   * const page = await sharePointService.getListItemsPaginated(
   *   'Empleados',
   *   1,
   *   10
   * );
   */
  getListItemsPaginated(listName: string, pageNumber: number, pageSize: number): Promise<{ items: IDataItem[]; totalCount: number }>;
}
