/**
 * Archivo: HelloWorld.tsx
 * Descripción: Componente React principal para el Web Part HelloWorld.
 * 
 * ARQUITECTURA CORRECTA - SEPARACIÓN DE RESPONSABILIDADES:
 * ========================================================
 * Este archivo contiene SOLO la lógica del componente y el JSX.
 * 
 * Las importaciones están organizadas en 4 categorías:
 * 1. React y utilidades de Microsoft
 * 2. Estilos (CSS Modules)
 * 3. Interfaces y tipos (Props, State, Data Types)
 * 4. Componentes de Fluent UI
 */

import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorld.module.scss';

// ============================================================================
// IMPORTACIONES DE INTERFACES Y TIPOS
// ============================================================================
import { IHelloWorldProps } from './interfaces/IHelloWorldProps';
import { IHelloWorldState } from './interfaces/IHelloWorldState';
import {
  IDataItem,
  ESTADOS_DISPONIBLES,
  DEPARTAMENTOS_DISPONIBLES,
} from './interfaces/IDataTypes';

// ============================================================================
// IMPORTACIONES DE FLUENT UI
// ============================================================================
import {
  PrimaryButton,
  Modal,
  Stack,
  Text,
  IconButton,
  DetailsList,
  IColumn,
  IGroup,
  SelectionMode,
  IStackTokens,
  IDetailsGroupDividerProps // <-- AÑADIDO: Interfaz correcta para evitar el uso de 'any'
} from '@fluentui/react';

// ============================================================================
// COMPONENTE PRINCIPAL
// ============================================================================
export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  constructor(props: IHelloWorldProps) {
    super(props);

    const dataItems: IDataItem[] = [
      {
        key: '1',
        id: 1,
        nombre: 'Juan García',
        departamento: DEPARTAMENTOS_DISPONIBLES.VENTAS,
        email: 'juan.garcia@empresa.com',
        estado: ESTADOS_DISPONIBLES.ACTIVO,
        fechaRegistro: '2026-07-13',
      },
      {
        key: '2',
        id: 2,
        nombre: 'María López',
        departamento: DEPARTAMENTOS_DISPONIBLES.VENTAS,
        email: 'maria.lopez@empresa.com',
        estado: ESTADOS_DISPONIBLES.ACTIVO,
        fechaRegistro: '2026-07-13', // Corregido un typo en el año de tu código original
      },
      {
        key: '3',
        id: 3,
        nombre: 'Carlos Rodríguez',
        departamento: DEPARTAMENTOS_DISPONIBLES.IT,
        email: 'carlos.rodriguez@empresa.com',
        estado: ESTADOS_DISPONIBLES.ACTIVO,
        fechaRegistro: '2026-07-14',
      },
      {
        key: '4',
        id: 4,
        nombre: 'Ana Martínez',
        departamento: DEPARTAMENTOS_DISPONIBLES.IT,
        email: 'ana.martinez@empresa.com',
        estado: ESTADOS_DISPONIBLES.INACTIVO,
        fechaRegistro: '2026-07-14',
      },
      {
        key: '5',
        id: 5,
        nombre: 'Pedro Sánchez',
        departamento: DEPARTAMENTOS_DISPONIBLES.RECURSOS_HUMANOS,
        email: 'pedro.sanchez@empresa.com',
        estado: ESTADOS_DISPONIBLES.ACTIVO,
        fechaRegistro: '2026-07-15',
      },
      {
        key: '6',
        id: 6,
        nombre: 'Laura Fernández',
        departamento: DEPARTAMENTOS_DISPONIBLES.RECURSOS_HUMANOS,
        email: 'laura.fernandez@empresa.com',
        estado: ESTADOS_DISPONIBLES.ACTIVO,
        fechaRegistro: '2026-07-15',
      },
    ];

    const columns: IColumn[] = [
      {
        key: 'id',
        name: 'ID',
        fieldName: 'id',
        minWidth: 50,
        maxWidth: 80,
        isResizable: true,
      },
      {
        key: 'nombre',
        name: 'Nombre',
        fieldName: 'nombre',
        minWidth: 150,
        isResizable: true,
      },
      {
        key: 'email',
        name: 'Email',
        fieldName: 'email',
        minWidth: 180,
        isResizable: true,
      },
      {
        key: 'estado',
        name: 'Estado',
        fieldName: 'estado',
        minWidth: 60,
        maxWidth: 70,
        isResizable: true,
        onRender: (item: IDataItem) => {
          const statusClass =
            item.estado === ESTADOS_DISPONIBLES.ACTIVO
              ? styles.statusActive
              : item.estado === ESTADOS_DISPONIBLES.INACTIVO
                ? styles.statusInactive
                : styles.statusPending;

          return <span className={statusClass}>{item.estado}</span>;
        },
      },
      {
        key: 'fechaRegistro',
        name: 'Fecha de Registro',
        fieldName: 'fechaRegistro',
        minWidth: 110,
        isResizable: true,
        onRender: (item: IDataItem) => {
          return (
            <div className={styles.centeredColumn}>
              {item.fechaRegistro}
            </div>
          );
        },
        headerClassName: styles.centeredColumnHeader,
      },
    ];

    const groups: IGroup[] = [
      {
        key: DEPARTAMENTOS_DISPONIBLES.VENTAS,
        name: DEPARTAMENTOS_DISPONIBLES.VENTAS,
        startIndex: 0,
        count: 2,
        level: 0,
        isCollapsed: false,
      },
      {
        key: DEPARTAMENTOS_DISPONIBLES.IT,
        name: DEPARTAMENTOS_DISPONIBLES.IT,
        startIndex: 2,
        count: 2,
        level: 0,
        isCollapsed: false,
      },
      {
        key: DEPARTAMENTOS_DISPONIBLES.RECURSOS_HUMANOS,
        name: DEPARTAMENTOS_DISPONIBLES.RECURSOS_HUMANOS,
        startIndex: 4,
        count: 2,
        level: 0,
        isCollapsed: false,
      },
    ];

    this.state = {
      isModalOpen: false,
      items: dataItems,
      groups: groups,
      columns: columns,
    };
  }

  private _abrirModal = (): void => {
    this.setState({ isModalOpen: true });
  };

  private _cerrarModal = (): void => {
    this.setState({ isModalOpen: false });
  };

  /**
   * CORRECCIÓN: Se utiliza un parámetro opcional para coincidir con la firma
   * esperada por Fluent UI y evitar el fallo de compilación cuando el valor
   * puede ser undefined.
   */
  private _onRenderGroupHeader = (props?: IDetailsGroupDividerProps): React.ReactElement | null => {
    if (!props || !props.group) {
        return null;
    }

    return (
      <div
        className={styles.groupHeader}
        onClick={() => props.onToggleCollapse!(props.group!)}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            props.onToggleCollapse!(props.group!);
          }
        }}
      >
        <IconButton
          iconProps={{
            iconName: props.group.isCollapsed ? 'ChevronRight' : 'ChevronDown',
          }}
          className={styles.groupHeaderIcon}
          aria-label={`${props.group.isCollapsed ? 'Expandir' : 'Contraer'} ${props.group.name}`}
        />
        <Text variant="mediumPlus" className={styles.groupHeaderText}>
          {props.group.name} ({props.group.count})
        </Text>
      </div>
    );
  };

  public render(): React.ReactElement<IHelloWorldProps> {
    const { isDarkTheme, hasTeamsContext, userDisplayName } = this.props;
    const { isModalOpen, items, groups, columns } = this.state;

    const sectionStackTokens: IStackTokens = { childrenGap: 20 };
    const modalStackTokens: IStackTokens = { childrenGap: 15 };
    const tableStackTokens: IStackTokens = { childrenGap: 10 };

    return (
      <section
        className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.darkTheme : ''}`}
      >
        <Stack tokens={sectionStackTokens}>
          <div className={styles.welcome}>
            <img
              alt="Imagen de bienvenida"
              src={
                isDarkTheme
                  ? require('../assets/welcome-dark.png')
                  : require('../assets/welcome-light.png')
              }
              className={styles.welcomeImage}
            />
            <Text
              variant="xxLarge"
              block
              styles={{ root: { fontWeight: '600', marginTop: '20px' } }}
            >
              Bienvenido, {escape(userDisplayName)}!
            </Text>
          </div>

          <Stack tokens={sectionStackTokens}>
            <Text variant="xLarge" styles={{ root: { fontWeight: '600' } }}>
              Tabla Avanzada con Agrupación
            </Text>
            <Text variant="medium">
              Este ejemplo muestra cómo implementar una tabla avanzada (DetailsList) con agrupación
              por departamento. Los datos se organizan automáticamente en grupos colapsables.
            </Text>
            <div>
              <PrimaryButton text="Abrir Información" onClick={this._abrirModal} />
            </div>
          </Stack>

          <Stack tokens={tableStackTokens} className={styles.detailsListContainer}>
            <Text variant="large" styles={{ root: { fontWeight: '600' } }}>
              Listado de Empleados por Departamento
            </Text>
            <DetailsList
              items={items}
              columns={columns}
              groups={groups}
              groupProps={{
                onRenderHeader: this._onRenderGroupHeader,
              }}
              selectionMode={SelectionMode.multiple}
              isHeaderVisible={true}
              compact={false}
              styles={{
                root: {
                  border: '1px solid #E1E1E1',
                  borderRadius: '4px',
                  overflow: 'hidden',
                },
              }}
            />
          </Stack>

          <Stack tokens={sectionStackTokens}>
            <Text variant="large" styles={{ root: { fontWeight: '600' } }}>
              Características Implementadas
            </Text>
            <ul style={{ margin: '0', paddingLeft: '20px' }}>
              <li>
                <Text variant="medium">
                  <strong>Tabla Avanzada:</strong> Utiliza DetailsList de Fluent UI con soporte
                  para múltiples columnas redimensionables
                </Text>
              </li>
              <li>
                <Text variant="medium">
                  <strong>Agrupación:</strong> Los datos se agrupan automáticamente por
                  departamento con capacidad de expandir/contraer
                </Text>
              </li>
              <li>
                <Text variant="medium">
                  <strong>Renderizado Personalizado:</strong> El estado se muestra con colores
                  personalizados (verde para Activo, rojo para Inactivo)
                </Text>
              </li>
              <li>
                <Text variant="medium">
                  <strong>Selección Múltiple:</strong> Permite seleccionar múltiples filas
                </Text>
              </li>
              <li>
                <Text variant="medium">
                  <strong>Columnas Redimensionables:</strong> Cada columna puede ajustarse
                  manualmente
                </Text>
              </li>
              <li>
                <Text variant="medium">
                  <strong>Arquitectura Profesional:</strong> Separación completa de
                  responsabilidades (Props, State, Types, Styles, Lógica)
                </Text>
              </li>
            </ul>
          </Stack>
        </Stack>

        <Modal isOpen={isModalOpen} onDismiss={this._cerrarModal} isBlocking={false}>
          <Stack tokens={modalStackTokens} styles={{ root: { padding: '24px', minWidth: '400px' } }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="xLarge" styles={{ root: { fontWeight: '600' } }}>
                Información de la Tabla
              </Text>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                ariaLabel="Cerrar modal"
                onClick={this._cerrarModal}
              />
            </Stack>

            <Text variant="medium">
              Esta tabla está construida con <strong>DetailsList</strong> de Fluent UI, que
              proporciona una forma robusta y accesible de mostrar datos tabulares con
              características avanzadas como agrupación, selección y renderizado personalizado.
            </Text>

            <Text variant="medium">
              Los datos se organizan en grupos por departamento, permitiendo a los usuarios
              expandir o contraer cada sección según sea necesario.
            </Text>

            <Text variant="medium" styles={{ root: { fontStyle: 'italic', color: '#605e5c' } }}>
              <strong>Nota Arquitectónica:</strong> Este componente demuestra las mejores prácticas
              de separación de responsabilidades, con Props, State, Types, Estilos y Lógica en
              archivos separados para máxima mantenibilidad y escalabilidad.
            </Text>

            <Stack horizontal horizontalAlign="end">
              <PrimaryButton text="Cerrar" onClick={this._cerrarModal} />
            </Stack>
          </Stack>
        </Modal>
      </section>
    );
  }
}