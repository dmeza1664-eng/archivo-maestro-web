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

CREATE TABLE IF NOT EXISTS usuarios (
  id INT UNSIGNED NOT NULL AUTO_INCREMENT,
  usuario VARCHAR(80) NOT NULL,
  nombre VARCHAR(160) NOT NULL,
  password_hash CHAR(128) NOT NULL,
  password_salt CHAR(32) NOT NULL,
  rol ENUM('admin', 'operador', 'consulta') NOT NULL DEFAULT 'operador',
  activo TINYINT(1) NOT NULL DEFAULT 1,
  ultimo_acceso TIMESTAMP NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_usuarios_usuario (usuario)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS sesiones (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  usuario_id INT UNSIGNED NOT NULL,
  token_hash CHAR(64) NOT NULL,
  expira_at TIMESTAMP NOT NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_sesiones_token_hash (token_hash),
  KEY ix_sesiones_usuario_id (usuario_id),
  KEY ix_sesiones_expira_at (expira_at),
  CONSTRAINT fk_sesiones_usuario
    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
    ON UPDATE CASCADE ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS respaldos_datos (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  tipo VARCHAR(80) NOT NULL,
  version INT UNSIGNED NOT NULL,
  periodo VARCHAR(20) NOT NULL DEFAULT 'global',
  archivos JSON NULL,
  contenido LONGTEXT NOT NULL,
  usuario_id INT UNSIGNED NOT NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_respaldos_tipo_periodo_version (tipo, periodo, version),
  KEY ix_respaldos_tipo_periodo_fecha (tipo, periodo, created_at),
  KEY ix_respaldos_usuario_id (usuario_id),
  CONSTRAINT fk_respaldos_usuario
    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS bitacora (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  usuario_id INT UNSIGNED NULL,
  accion VARCHAR(80) NOT NULL,
  entidad VARCHAR(80) NOT NULL,
  clave_entidad VARCHAR(160) NOT NULL DEFAULT '',
  detalle JSON NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY ix_bitacora_fecha (created_at),
  KEY ix_bitacora_usuario_id (usuario_id),
  KEY ix_bitacora_entidad (entidad, clave_entidad),
  CONSTRAINT fk_bitacora_usuario
    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
    ON UPDATE CASCADE ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS importaciones_ventas (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  usuario_id INT UNSIGNED NOT NULL,
  archivo VARCHAR(255) NOT NULL DEFAULT '',
  filas_recibidas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_validas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_rechazadas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_consolidadas INT UNSIGNED NOT NULL DEFAULT 0,
  registros_insertados INT UNSIGNED NOT NULL DEFAULT 0,
  registros_actualizados INT UNSIGNED NOT NULL DEFAULT 0,
  fecha_min DATE NULL,
  fecha_max DATE NULL,
  detalle JSON NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY ix_importaciones_ventas_fecha (created_at),
  KEY ix_importaciones_ventas_usuario (usuario_id),
  CONSTRAINT fk_importaciones_ventas_usuario
    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS bajas_diarias (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  fecha DATE NOT NULL,
  producto_id INT UNSIGNED NOT NULL,
  cantidad DECIMAL(14,4) NOT NULL DEFAULT 0,
  sucursal VARCHAR(120) NOT NULL DEFAULT '',
  motivo VARCHAR(160) NOT NULL DEFAULT '',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY ux_bajas_fecha_producto_sucursal_motivo (fecha, producto_id, sucursal, motivo),
  KEY ix_bajas_fecha (fecha),
  KEY ix_bajas_producto_id (producto_id),
  CONSTRAINT fk_bajas_producto
    FOREIGN KEY (producto_id) REFERENCES productos (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS importaciones_operativas (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  tipo VARCHAR(40) NOT NULL,
  usuario_id INT UNSIGNED NOT NULL,
  archivo VARCHAR(255) NOT NULL DEFAULT '',
  filas_recibidas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_validas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_rechazadas INT UNSIGNED NOT NULL DEFAULT 0,
  filas_consolidadas INT UNSIGNED NOT NULL DEFAULT 0,
  registros_insertados INT UNSIGNED NOT NULL DEFAULT 0,
  registros_actualizados INT UNSIGNED NOT NULL DEFAULT 0,
  fecha_min DATE NULL,
  fecha_max DATE NULL,
  detalle JSON NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY ix_importaciones_operativas_tipo_fecha (tipo, created_at),
  KEY ix_importaciones_operativas_usuario (usuario_id),
  CONSTRAINT fk_importaciones_operativas_usuario
    FOREIGN KEY (usuario_id) REFERENCES usuarios (id)
    ON UPDATE CASCADE ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
