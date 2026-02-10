/**
 * Encuentra el mejor cliente coincidente usando regex y comparación de palabras
 * @param {string} clienteBackend - Nombre del cliente que viene del backend/IA
 * @param {Array} clientsList - Array de objetos cliente con propiedad 'name'
 * @returns {string|null} - El nombre del cliente más coincidente o null si no hay coincidencias suficientes
 */
export function encontrarClienteCoincidente(clienteBackend, clientsList) {
  // Validar parámetros
  if (!clienteBackend || typeof clienteBackend !== 'string') {
    return null;
  }
  
  if (!clientsList || !Array.isArray(clientsList) || clientsList.length === 0) {
    return null;
  }
  
  const clienteNormalizado = clienteBackend.trim();
  if (clienteNormalizado.length === 0) {
    return null;
  }
  
  // Paso 1: Coincidencia exacta (case-insensitive)
  const clienteLower = clienteNormalizado.toLowerCase();
  for (const cliente of clientsList) {
    if (cliente && cliente.name && typeof cliente.name === 'string') {
      const clienteNameLower = cliente.name.trim().toLowerCase();
      if (clienteNameLower === clienteLower) {
        return cliente.name; // Retornar el nombre original del cliente
      }
    }
  }
  
  // Paso 2: Comparación por palabras con regex
  const clienteUpper = clienteNormalizado.toUpperCase();
  // Dividir en palabras usando regex (espacios, guiones, etc.)
  const palabrasBackend = clienteUpper.split(/\s+|[-–—]/).filter(p => p.length > 0);
  
  if (palabrasBackend.length === 0) {
    return null;
  }
  
  let mejorCoincidencia = null;
  let maxCoincidencias = 0;
  let mejorPuntuacion = 0;
  
  // Paso 3: Manejo de variaciones sin espacios
  // Normalizar removiendo espacios para comparación
  const clienteSinEspacios = clienteUpper.replace(/\s+/g, '');
  
  for (const cliente of clientsList) {
    if (!cliente || !cliente.name || typeof cliente.name !== 'string') {
      continue;
    }
    
    const clienteName = cliente.name.trim();
    const clienteNameUpper = clienteName.toUpperCase();
    const clienteNameSinEspacios = clienteNameUpper.replace(/\s+/g, '');
    
    // Verificar coincidencia sin espacios (bidireccional)
    let tieneCoincidenciaSinEspacios = false;
    if (clienteSinEspacios.length > 0 && clienteNameSinEspacios.length > 0) {
      if (clienteSinEspacios === clienteNameSinEspacios) {
        // Coincidencia exacta sin espacios - retornar inmediatamente
        return cliente.name;
      }
      // Verificar includes bidireccional
      if (clienteSinEspacios.includes(clienteNameSinEspacios) || 
          clienteNameSinEspacios.includes(clienteSinEspacios)) {
        tieneCoincidenciaSinEspacios = true;
      }
    }
    
    // Dividir nombre del cliente en palabras
    const palabrasLista = clienteNameUpper.split(/\s+|[-–—]/).filter(p => p.length > 0);
    
    if (palabrasLista.length === 0) {
      continue;
    }
    
    // Contar palabras que coinciden
    let coincidencias = 0;
    let puntuacion = 0;
    
    palabrasBackend.forEach(palabraBackend => {
      // Buscar la palabra en la lista (coincidencia exacta o parcial)
      const encontrado = palabrasLista.some(palabraLista => {
        // Coincidencia exacta
        if (palabraLista === palabraBackend) {
          puntuacion += 2;
          return true;
        }
        // Coincidencia parcial (una contiene a la otra)
        if (palabraLista.includes(palabraBackend) || palabraBackend.includes(palabraLista)) {
          // Solo contar si ambas palabras tienen al menos 3 caracteres para evitar falsos positivos
          if (palabraBackend.length >= 3 && palabraLista.length >= 3) {
            puntuacion += 1;
            return true;
          }
        }
        return false;
      });
      
      if (encontrado) {
        coincidencias++;
      }
    });
    
    // Agregar puntuación por coincidencia sin espacios
    if (tieneCoincidenciaSinEspacios) {
      puntuacion += 1.5;
    }
    
    // Si hay más de una palabra coincidente, considerar esta opción
    // O si hay una coincidencia muy significativa (alta puntuación)
    if (coincidencias > maxCoincidencias && (coincidencias > 1 || puntuacion >= 2)) {
      maxCoincidencias = coincidencias;
      mejorPuntuacion = puntuacion;
      mejorCoincidencia = cliente.name;
    } else if (coincidencias === maxCoincidencias && puntuacion > mejorPuntuacion) {
      // Si hay el mismo número de coincidencias, elegir el de mayor puntuación
      mejorPuntuacion = puntuacion;
      mejorCoincidencia = cliente.name;
    }
  }
  
  // Paso 4: Retornar mejor coincidencia si supera el umbral mínimo
  // Mínimo requerido: 2 puntos o al menos 2 palabras coincidentes
  if (maxCoincidencias >= 2 || mejorPuntuacion >= 2) {
    return mejorCoincidencia;
  }
  
  return null;
}
