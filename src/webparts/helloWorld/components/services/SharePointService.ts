/**
 * Archivo: SharePointService.ts
 * Descripción: Implementación del servicio para interactuar con SharePoint.
 * 
 * Este servicio encapsula toda la lógica de comunicación con SharePoint,
 * permitiendo que el componente React no necesite conocer los detalles
 * de las llamadas REST.
 * 
 * Patrones aplicados:
 * - Service Pattern: Centraliza la lógica de datos
 * - Dependency Injection: Recibe SPHttpClient en el constructor
 * - Promise-based: Usa async/await para mejor legibilidad
 * 
 * Autor: Manus AI
 * Fecha: 2026
 */

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// Interfaz para los datos procesados que usará tu aplicación (React, etc.)
export interface IDataItem {
  key: string;
  id: number;
  nombre: string;
  departamento: string;
  email: string;
  estado: string;
  fechaRegistro: string;
}

export interface IDataResponse {
  items: IDataItem[];
  totalCount: number;
  hasMorePages: boolean;
}

// Interfaz cruda que describe exactamente lo que devuelve la API de SharePoint
interface ISPListItem {
  ID: number;
  Nombre: string;
  Departamento: string;
  Email: string;
  Estado: string;
  FechaRegistro: string;
}

export interface ISharePointService {
  getListItems(listName: string): Promise<IDataItem[]>;
  getListItemsFiltered(listName: string, filterQuery: string): Promise<IDataItem[]>;
  getListItemsPaginated(listName: string, pageNumber: number, pageSize: number): Promise<IDataResponse>;
  getListItemById(listName: string, itemId: number): Promise<IDataItem>;
}

export class SharePointService implements ISharePointService {
  private spHttpClient: SPHttpClient;
  private pageContext: WebPartContext['pageContext'];

  constructor(context: WebPartContext) {
    this.spHttpClient = context.spHttpClient;
    this.pageContext = context.pageContext;
  }

  public async getListItems(listName: string): Promise<IDataItem[]> {
    try {
      const url = `${this.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url, 
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error al obtener elementos: ${response.statusText}`);
      }

      const data = await response.json();

      // Mapeamos los datos usando la interfaz cruda ISPListItem
      const items: IDataItem[] = data.value.map((item: ISPListItem) => ({
        key: item.ID.toString(),
        id: item.ID,
        nombre: item.Nombre || '',
        departamento: item.Departamento || '',
        email: item.Email || '',
        estado: item.Estado || 'Pendiente',
        fechaRegistro: item.FechaRegistro ? item.FechaRegistro.split('T')[0] : '',
      }));

      return items;
    } catch (error) {
      console.error('Error en getListItems:', error);
      throw error;
    }
  }

  public async getListItemsFiltered(listName: string, filterQuery: string): Promise<IDataItem[]> {
    try {
      const url = `${this.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items?$filter=${filterQuery}`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url, 
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error al obtener elementos filtrados: ${response.statusText}`);
      }

      const data = await response.json();

      // Mapeamos los datos usando la interfaz cruda ISPListItem
      const items: IDataItem[] = data.value.map((item: ISPListItem) => ({
        key: item.ID.toString(),
        id: item.ID,
        nombre: item.Nombre || '',
        departamento: item.Departamento || '',
        email: item.Email || '',
        estado: item.Estado || 'Pendiente',
        fechaRegistro: item.FechaRegistro ? item.FechaRegistro.split('T')[0] : '',
      }));

      return items;
    } catch (error) {
      console.error('Error en getListItemsFiltered:', error);
      throw error;
    }
  }

  public async getListItemsPaginated(
    listName: string,
    pageNumber: number,
    pageSize: number
  ): Promise<IDataResponse> {
    try {
      // Lógica básica de paginación con SharePoint REST API (top y skip)
      const skipValue = (pageNumber - 1) * pageSize;
      const url = `${this.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items?$top=${pageSize}&$skip=${skipValue}`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url, 
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error al obtener elementos paginados: ${response.statusText}`);
      }

      const data = await response.json();

      // Mapeamos los datos usando la interfaz cruda ISPListItem
      const items: IDataItem[] = data.value.map((item: ISPListItem) => ({
        key: item.ID.toString(),
        id: item.ID,
        nombre: item.Nombre || '',
        departamento: item.Departamento || '',
        email: item.Email || '',
        estado: item.Estado || 'Pendiente',
        fechaRegistro: item.FechaRegistro ? item.FechaRegistro.split('T')[0] : '',
      }));

      return {
        items,
        totalCount: items.length, // Nota: Para un totalCount real en SP, suele requerir otra llamada o lógica adicional
        hasMorePages: !!data['@odata.nextLink'] 
      };
    } catch (error) {
      console.error('Error en getListItemsPaginated:', error);
      throw error;
    }
  }

  public async getListItemById(listName: string, itemId: number): Promise<IDataItem> {
    try {
      const url = `${this.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})`;

      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error al obtener elemento por ID: ${response.status}`);
      }

      // Tipamos directamente la respuesta JSON a nuestra interfaz
      const item: ISPListItem = await response.json();

      return {
        key: item.ID.toString(),
        id: item.ID,
        nombre: item.Nombre || '',
        departamento: item.Departamento || '',
        email: item.Email || '',
        estado: item.Estado || 'Pendiente',
        fechaRegistro: item.FechaRegistro ? item.FechaRegistro.split('T')[0] : '',
      };
    } catch (error) {
      console.error('Error en getListItemById:', error);
      throw error;
    }
  }
}