/**
 * Archivo: IHelloWorldProps.ts
 * Descripción: Define las propiedades que se pasan al componente HelloWorld desde el Web Part.
 * Responsabilidad única: Definir la interfaz de propiedades del componente.

 * Interfaz para las propiedades del componente HelloWorld
 * Estas propiedades se pasan desde el Web Part principal (HelloWorldWebPart.ts)
 * hacia el componente React.
 * 
 * Nota: Las propiedades son de solo lectura (readonly) y no deben cambiar durante
 * la vida útil del componente. Si necesitas datos dinámicos, usa el estado (State).
*/
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHelloWorldProps {
  description: string;        // Descripción del Web Part (generalmente desde las propiedades configurables)
  isDarkTheme: boolean;       // Indica si el tema actual es oscuro (dark mode), Utilizado para cambiar la imagen de bienvenida y otros estilos condicionales
  environmentMessage: string; //  Mensaje de entorno (ej. "Ejecutándose en SharePoint Online"), Útil para mostrar información de contexto
  hasTeamsContext: boolean;   // Indica si el Web Part se está ejecutando dentro de Microsoft Teams, Esto puede afectar la apariencia y el comportamiento del componente
  userDisplayName: string;    // Nombre para mostrar del usuario actual, Se puede usar para personalizar la experiencia del usuario 
  context: WebPartContext;
  listName: string;
}
