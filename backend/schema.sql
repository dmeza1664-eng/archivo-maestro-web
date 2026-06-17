CREATE TABLE IF NOT EXISTS productos (
  id INT UNSIGNED NOT NULL AUTO_INCREMENT,
  codigo VARCHAR(120) NOT NULL,
  nombre VARCHAR(255) NOT NULL,
  activo TINYINT(1) NOT NULL DEFAULT 1,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_productos_codigo (codigo)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS homologaciones_productos (
  id INT UNSIGNED NOT NULL AUTO_INCREMENT,
  codigo_origen VARCHAR(120) NOT NULL,
  nombre_origen VARCHAR(255) NULL,
  producto_id INT UNSIGNED NOT NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_homologacion_codigo_origen (codigo_origen),
  KEY ix_homologaciones_producto_id (producto_id),
  CONSTRAINT fk_homologaciones_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS ventas_diarias (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  fecha DATE NOT NULL,
  producto_id INT UNSIGNED NOT NULL,
  cantidad DECIMAL(14,4) NOT NULL DEFAULT 0,
  importe DECIMAL(14,2) NULL,
  canal VARCHAR(120) NOT NULL DEFAULT '',
  cliente VARCHAR(180) NOT NULL DEFAULT '',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_ventas_fecha_producto_canal_cliente (fecha, producto_id, canal, cliente),
  KEY ix_ventas_fecha (fecha),
  KEY ix_ventas_producto_id (producto_id),
  CONSTRAINT fk_ventas_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS stock_fijo (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  mes CHAR(7) NOT NULL,
  producto_id INT UNSIGNED NOT NULL,
  cantidad DECIMAL(14,4) NOT NULL DEFAULT 0,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_stock_mes_producto (mes, producto_id),
  KEY ix_stock_producto_id (producto_id),
  CONSTRAINT fk_stock_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS produccion_real (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  fecha DATE NOT NULL,
  producto_id INT UNSIGNED NOT NULL,
  cantidad DECIMAL(14,4) NOT NULL DEFAULT 0,
  turno VARCHAR(80) NOT NULL DEFAULT '',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_produccion_fecha_producto_turno (fecha, producto_id, turno),
  KEY ix_produccion_fecha (fecha),
  KEY ix_produccion_producto_id (producto_id),
  CONSTRAINT fk_produccion_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS pronostico_diario (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  fecha DATE NOT NULL,
  producto_id INT UNSIGNED NOT NULL,
  cantidad_pronosticada DECIMAL(14,4) NOT NULL DEFAULT 0,
  metodo VARCHAR(80) NOT NULL DEFAULT 'promedio_ventas',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_pronostico_fecha_producto_metodo (fecha, producto_id, metodo),
  KEY ix_pronostico_fecha (fecha),
  KEY ix_pronostico_producto_id (producto_id),
  CONSTRAINT fk_pronostico_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
