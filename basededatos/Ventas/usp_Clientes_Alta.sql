-- =============================================
-- Stored Procedure: usp_Clientes_Alta
-- Descripción: Inserta una nueva solicitud de alta de cliente
-- Base de datos: ventas
-- Tabla: dbo.Clientes_Alta
-- =============================================

USE [ventas]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_Clientes_Alta]') AND type in (N'P', N'PC'))
    DROP PROCEDURE [dbo].[usp_Clientes_Alta]
GO

CREATE PROCEDURE [dbo].[usp_Clientes_Alta]
    @VendedorID VARCHAR(20),
    @VendedorNombre VARCHAR(120),
    @Nombre VARCHAR(80),
    @Apellido VARCHAR(80),
    @Direccion VARCHAR(160),
    @CUITCUIL VARCHAR(20),
    @Provincia VARCHAR(80),
    @Ciudad VARCHAR(80),
    @ConstanciaAfipArchivo VARCHAR(255),
    @ConstanciaAfipRuta VARCHAR(500),
    @AltaClienteID INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    -- Validar que no exista una solicitud pendiente para el mismo CUIT
    IF EXISTS (
        SELECT 1 FROM dbo.Clientes_Alta
        WHERE CUITCUIL = @CUITCUIL AND Status = 0
    )
    BEGIN
        RAISERROR('Ya existe una solicitud pendiente para este CUIT/CUIL.', 16, 1);
        RETURN;
    END

    -- Insertar la nueva solicitud
    INSERT INTO dbo.Clientes_Alta (
        VendedorID,
        VendedorNombre,
        Nombre,
        Apellido,
        Direccion,
        CUITCUIL,
        Provincia,
        Ciudad,
        ConstanciaAfipArchivo,
        ConstanciaAfipRuta,
        Status,
        FechaSolicitud,
        UsuarioAlta
    ) VALUES (
        @VendedorID,
        @VendedorNombre,
        @Nombre,
        @Apellido,
        @Direccion,
        @CUITCUIL,
        @Provincia,
        @Ciudad,
        @ConstanciaAfipArchivo,
        @ConstanciaAfipRuta,
        0, -- Status: 0 = Pendiente
        GETDATE(), -- FechaSolicitud
        @VendedorID -- UsuarioAlta
    );

    -- Obtener el ID generado
    SET @AltaClienteID = SCOPE_IDENTITY();

    -- Log de auditoría (opcional)
    INSERT INTO dbo.Clientes_Alta_Log (
        AltaClienteID,
        Accion,
        Usuario,
        Fecha
    ) VALUES (
        @AltaClienteID,
        'CREADO',
        @VendedorID,
        GETDATE()
    );

END
GO

-- =============================================
-- Tabla: dbo.Clientes_Alta (si no existe)
-- =============================================

IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Clientes_Alta' AND xtype='U')
BEGIN
    CREATE TABLE dbo.Clientes_Alta (
        AltaClienteID INT IDENTITY(1,1) PRIMARY KEY,
        VendedorID VARCHAR(20) NOT NULL,
        VendedorNombre VARCHAR(120) NOT NULL,
        Nombre VARCHAR(80) NOT NULL,
        Apellido VARCHAR(80) NOT NULL,
        Direccion VARCHAR(160) NOT NULL,
        CUITCUIL VARCHAR(20) NOT NULL,
        Provincia VARCHAR(80) NOT NULL,
        Ciudad VARCHAR(80) NOT NULL,
        ConstanciaAfipArchivo VARCHAR(255) NOT NULL,
        ConstanciaAfipRuta VARCHAR(500) NOT NULL,
        Status TINYINT NOT NULL DEFAULT 0, -- 0=Pendiente, 1=Procesado
        FechaSolicitud DATETIME NOT NULL DEFAULT GETDATE(),
        FechaProcesado DATETIME NULL,
        UsuarioAlta VARCHAR(20) NOT NULL,
        UsuarioProcesado VARCHAR(20) NULL,
        FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
        FechaModificacion DATETIME NULL
    );

    -- Índices para mejor performance
    CREATE INDEX IX_Clientes_Alta_Status ON dbo.Clientes_Alta (Status);
    CREATE INDEX IX_Clientes_Alta_CUITCUIL ON dbo.Clientes_Alta (CUITCUIL);
    CREATE INDEX IX_Clientes_Alta_VendedorID ON dbo.Clientes_Alta (VendedorID);
END
GO

-- =============================================
-- Tabla de Log: dbo.Clientes_Alta_Log (opcional)
-- =============================================

IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Clientes_Alta_Log' AND xtype='U')
BEGIN
    CREATE TABLE dbo.Clientes_Alta_Log (
        LogID INT IDENTITY(1,1) PRIMARY KEY,
        AltaClienteID INT NOT NULL,
        Accion VARCHAR(50) NOT NULL, -- CREADO, PROCESADO, etc.
        Usuario VARCHAR(20) NOT NULL,
        Fecha DATETIME NOT NULL DEFAULT GETDATE(),
        Detalles VARCHAR(500) NULL,
        FOREIGN KEY (AltaClienteID) REFERENCES dbo.Clientes_Alta(AltaClienteID)
    );
END
GO

-- =============================================
-- Stored Procedure para procesar solicitudes (opcional)
-- =============================================

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_Clientes_Alta_Procesar]') AND type in (N'P', N'PC'))
    DROP PROCEDURE [dbo].[usp_Clientes_Alta_Procesar]
GO

CREATE PROCEDURE [dbo].[usp_Clientes_Alta_Procesar]
    @AltaClienteID INT,
    @UsuarioProcesado VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;

    -- Validar que la solicitud exista y esté pendiente
    IF NOT EXISTS (
        SELECT 1 FROM dbo.Clientes_Alta
        WHERE AltaClienteID = @AltaClienteID AND Status = 0
    )
    BEGIN
        RAISERROR('La solicitud no existe o ya fue procesada.', 16, 1);
        RETURN;
    END

    -- Actualizar la solicitud como procesada
    UPDATE dbo.Clientes_Alta
    SET
        Status = 1,
        FechaProcesado = GETDATE(),
        UsuarioProcesado = @UsuarioProcesado,
        FechaModificacion = GETDATE()
    WHERE AltaClienteID = @AltaClienteID;

    -- Log de auditoría
    INSERT INTO dbo.Clientes_Alta_Log (
        AltaClienteID,
        Accion,
        Usuario,
        Fecha,
        Detalles
    ) VALUES (
        @AltaClienteID,
        'PROCESADO',
        @UsuarioProcesado,
        GETDATE(),
        'Solicitud procesada por usuario de administración'
    );

END
GO