/**
 * Archivo: HelloWorldWebPart.ts (ACTUALIZADO)
 * Descripción: Web Part principal que instancia el componente React con conexión a SharePoint.
 * 
 * Cambios respecto a la versión original:
 * - Se añade la propiedad 'listName' en las propiedades configurables
 * - Se pasa el contexto (this.context) al componente React
 * - Se pasa el nombre de la lista al componente
 * 
 /**
 * Archivo: HelloWorldWebPart.ts (ACTUALIZADO)
 * Descripción: Web Part principal que instancia el componente React con conexión a SharePoint.
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import HelloWorld from './components/HelloWorld';

export interface IHelloWorldWebPartProps {
  description: string;
  listName: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _hasTeamsContext: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      const teamsContext = this.context.microsoftTeams;
      if (teamsContext) {
        this._hasTeamsContext = true;
      }
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const isDarkTheme = this._isDarkTheme;

    const body = document.body;
    body.setAttribute('data-theme', isDarkTheme ? 'dark' : 'light');
  }

  public render(): void {
    const element: React.ReactElement = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._getEnvironmentMessage(),
        hasTeamsContext: this._hasTeamsContext,
        userDisplayName: this.context.pageContext.user.displayName || 'Usuario',
        context: this.context,
        listName: this.properties.listName || 'Empleados',
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.isServedFromLocalhost
        ? 'Ejecutándose en Microsoft Teams (localhost)'
        : 'Ejecutándose en Microsoft Teams';
    }

    return this.context.isServedFromLocalhost
      ? 'Ejecutándose en SharePoint Online (localhost)'
      : 'Ejecutándose en SharePoint Online';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configuración de la tabla avanzada',
          },
          groups: [
            {
              groupName: 'Configuración General',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Descripción',
                  value: this.properties.description,
                }),
                PropertyPaneTextField('listName', {
                  label: 'Nombre de la lista de SharePoint',
                  description: 'Especifica el nombre exacto de la lista (ej: Empleados, Contactos)',
                  value: this.properties.listName || 'Empleados',
                  placeholder: 'Empleados',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}