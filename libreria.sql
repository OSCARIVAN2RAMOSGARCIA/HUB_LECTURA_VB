IF NOT EXISTS (SELECT * FROM sys.schemas WHERE name = 'Biblioteca')
BEGIN
    EXEC('CREATE SCHEMA Biblioteca');
END
GO

-- Tabla Géneros
CREATE TABLE Biblioteca.Generos (
    id_genero INT IDENTITY(1,1) PRIMARY KEY,
    nombre_genero NVARCHAR(50) NOT NULL,
    descripcion NVARCHAR(255) NULL
);
GO

-- Tabla Libros (corregido el nombre de "Libros" en lugar de "Libros")
CREATE TABLE Biblioteca.Libros (
    id_libro INT IDENTITY(1,1) PRIMARY KEY,
    titulo NVARCHAR(100) NOT NULL,
    autor NVARCHAR(100) NOT NULL,
    id_genero INT NULL,
    anio_publicacion INT NULL,
    editorial NVARCHAR(50) NULL,
    disponible BIT DEFAULT 1,
    CONSTRAINT FK_Libros_Generos FOREIGN KEY (id_genero) REFERENCES Biblioteca.Generos(id_genero)
);
GO

-- Tabla Préstamos
CREATE TABLE Biblioteca.Prestamos (
    id_prestamo INT IDENTITY(1,1) PRIMARY KEY,
    id_libro INT NOT NULL,
    nombre_persona NVARCHAR(100) NOT NULL,
    fecha_prestamo DATE NOT NULL DEFAULT GETDATE(),
    fecha_devolucion DATE NULL,
    devuelto BIT DEFAULT 0,
    CONSTRAINT FK_Prestamos_Libros FOREIGN KEY (id_libro) REFERENCES Biblioteca.Libros(id_libro)
);
GO

-- Insertar datos de ejemplo
INSERT INTO Biblioteca.Generos (nombre_genero, descripcion) VALUES 
(N'Novela', N'Obras de ficción narrativa'),
(N'Ciencia Ficción', N'Literatura basada en supuestos logros científicos'),
(N'Historia', N'Libros sobre eventos históricos');

INSERT INTO Biblioteca.Libros (titulo, autor, id_genero, anio_publicacion, editorial) VALUES 
(N'Cien años de soledad', N'Gabriel García Márquez', 1, 1967, N'Sudamericana'),
(N'1984', N'George Orwell', 2, 1949, N'Secker & Warburg'),
(N'Breve historia del mundo', N'Ernst H. Gombrich', 3, 1935, N'Little, Brown');

INSERT INTO Biblioteca.Prestamos (id_libro, nombre_persona, fecha_prestamo, fecha_devolucion, devuelto) VALUES 
(1, N'María López', '2023-05-10', '2023-05-24', 1),
(2, N'Juan Pérez', '2023-06-01', NULL, 0);
GO