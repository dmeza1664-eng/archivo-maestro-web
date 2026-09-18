function validateCredentials(usuario, password) {
  if (!/^[A-Za-z0-9._-]{3,80}$/.test(usuario || '')) {
    const error = new Error('El usuario debe tener al menos 3 caracteres y usar letras, números, punto, guion o guion bajo');
    error.status = 400;
    throw error;
  }
  if (String(password || '').length < 4) {
    const error = new Error('La contraseña debe tener al menos 4 caracteres');
    error.status = 400;
    throw error;
  }
}

module.exports = { validateCredentials };
