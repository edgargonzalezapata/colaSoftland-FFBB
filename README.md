# Generador CSV Softland - Feria Bio Bio

Aplicación Windows Forms desarrollada para la Feria Bio Bio que genera archivos CSV compatibles con el sistema ERP Softland. Esta herramienta automatiza el procesamiento de diferentes tipos de transacciones financieras, incluyendo pagos anticipados, cheques y ventas.

## Características Principales

- Generación automática de archivos CSV compatibles con Softland
- Procesamiento de múltiples tipos de transacciones:
  - Pagos directos
  - Pagos anticipados
  - Pagos con cheques
  - Pagos de ventas
- Soporte para diferentes tipos de documentos:
  - Liquidación de Facturas (LF)
  - Facturas de Compra (FC)
  - Notas de Crédito
- Gestión de correlativos automática
- Validación de datos y manejo de errores robusto
- Interfaz de usuario intuitiva y fácil de usar

## Requisitos del Sistema

- Sistema Operativo: Windows
- .NET Framework 4.5 o superior
- Visual Studio 2019 o superior (para desarrollo)
- SQL Server 2012 o superior
- Acceso al sistema Softland ERP

## Configuración Inicial

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/edgargonzalezapata/colaSoftland-FFBB.git
   ```

2. Abrir la solución `colaSoftland.sln` en Visual Studio

3. Restaurar paquetes NuGet si es necesario

4. Configurar la cadena de conexión a la base de datos:
   - Ubicar el archivo `App.config`
   - Modificar la cadena de conexión con los datos de su servidor

5. Compilar la solución

6. Ejecutar la aplicación

## Uso

### Procesamiento de Transacciones

1. **Pagos Directos**
   - Seleccionar el tipo de documento
   - Ingresar los datos del pago
   - Generar el archivo CSV

2. **Pagos con Cheques**
   - Ingresar datos del cheque
   - Validar información bancaria
   - Procesar el pago

3. **Pagos Anticipados**
   - Registrar el anticipo
   - Asociar a documentos futuros
   - Generar comprobantes

4. **Exportación a Softland**
   - Verificar datos generados
   - Exportar archivo CSV
   - Validar en sistema Softland

## Estructura del Proyecto

- `Controller/`: Lógica de negocio y procesamiento
- `Model/`: Clases y modelos de datos
- `View/`: Formularios e interfaces de usuario
- `Conection/`: Manejo de conexiones a base de datos

## Mantenimiento

### Respaldo de Datos
- Realizar copias de seguridad periódicas de la base de datos
- Mantener respaldo de los archivos CSV generados

### Actualizaciones
- Verificar periódicamente nuevas versiones
- Actualizar según cambios en Softland

## Soporte

Para reportar problemas o solicitar nuevas funcionalidades:
1. Abrir un nuevo Issue en GitHub
2. Describir detalladamente el problema o sugerencia
3. Incluir capturas de pantalla si es necesario

## Contribuir

1. Hacer Fork del repositorio
2. Crear una rama para su funcionalidad (`git checkout -b feature/NuevaFuncionalidad`)
3. Hacer Commit de sus cambios (`git commit -m 'Agregar nueva funcionalidad'`)
4. Hacer Push a la rama (`git push origin feature/NuevaFuncionalidad`)
5. Crear un Pull Request

## Licencia

Este proyecto es software propietario de Feria Bio Bio. Todos los derechos reservados.

## Autores

- Edgar González - Desarrollo Principal
- Feria Bio Bio - Propietario 