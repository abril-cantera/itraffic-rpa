/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { crearReservaEnITraffic } from './rpaClient.js';
import { servicesList } from './servicesList.js';
import { encontrarClienteCoincidente } from './clienteMatcher.js';

// Estado global para datos maestros
const masterData = {
  sellers: [],
  clients: [],
  statuses: [],
  reservationTypes: [],
  genders: [],
  documentTypes: [],
  countries: [],
  loaded: false
};

// Estado global para la extracción actual
let extractionState = {
  didExtractionExist: false, // Si existe una extracción (para isReExtract)
  doesReservationExist: false, // Si existe una reserva (para determinar crear/editar)
  reservationCode: null,
  originData: null, // Data original de la extracción
  createdReservationCode: null // Código de reserva creada/editada
};

// Guardar HTML original de datosReservaSection para restaurarlo
let datosReservaSectionOriginalHTML = null;

// Función para mostrar mensajes al usuario
function mostrarMensaje(mensaje, tipo = "info") {
  // Crear elemento de mensaje
  const mensajeDiv = document.createElement("div");
  mensajeDiv.className = `status-message ${tipo}`;
  mensajeDiv.textContent = mensaje;
  mensajeDiv.style.position = "fixed";
  mensajeDiv.style.top = "20px";
  mensajeDiv.style.left = "50%";
  mensajeDiv.style.transform = "translateX(-50%)";
  mensajeDiv.style.zIndex = "10000";
  mensajeDiv.style.minWidth = "200px";
  mensajeDiv.style.maxWidth = "90%";
  mensajeDiv.style.textAlign = "center";
  mensajeDiv.style.animation = "slideDown 0.3s ease";
  
  document.body.appendChild(mensajeDiv);
  
  // Remover después de 3 segundos
  setTimeout(() => {
    mensajeDiv.style.animation = "slideUp 0.3s ease";
    setTimeout(() => {
      if (mensajeDiv.parentNode) {
        mensajeDiv.parentNode.removeChild(mensajeDiv);
      }
    }, 300);
  }, 3000);
}

Office.onReady((info) => {
  try {
  if (info.host === Office.HostType.Outlook) {
      // Guardar HTML original de datosReservaSection
      const datosReservaSection = document.getElementById("datosReservaSection");
      if (datosReservaSection) {
        datosReservaSectionOriginalHTML = datosReservaSection.innerHTML;
      }
      // Ocultar mensaje de sideload si existe
      const sideloadMsg = document.getElementById("sideload-msg");
      if (sideloadMsg) {
        sideloadMsg.style.display = "none";
      }
      
      // Mostrar el cuerpo de la aplicación
      const appBody = document.getElementById("app-body");
      if (appBody) {
        appBody.style.display = "flex";
      }
      
      // Controlar scroll inicial - no mostrar scroll cuando solo hay título y botón
      const mainContainer = document.querySelector(".ms-welcome__main");
      if (mainContainer) {
        mainContainer.classList.add("no-scroll");
      }
      
      // Asignar evento al botón de extraer
      const runButton = document.getElementById("run");
      if (runButton) {
        runButton.onclick = function() {
          try {
            run();
          } catch (error) {
            mostrarMensaje("Error al extraer datos: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de re-extraer
      const reextractButton = document.getElementById("reextract");
      if (reextractButton) {
        // Ocultar inicialmente (solo se muestra cuando hay resultados)
        reextractButton.style.display = "none";
        
        reextractButton.onclick = function() {
          try {
            // Ocultar resultados y volver a extraer
            const resultsDiv = document.getElementById("results");
            resultsDiv.style.display = "none";
            // Ocultar botón de re-extracción mientras se procesa
            reextractButton.style.display = "none";
            // Llamar a run con isReExtract = true
            run(true);
          } catch (error) {
            mostrarMensaje("Error al re-extraer datos: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de guardar
      const guardarButton = document.getElementById("guardar");
      if (guardarButton) {
        guardarButton.onclick = function() {
          try {
            guardarDatos();
          } catch (error) {
            mostrarMensaje("Error al guardar datos: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de agregar pasajero
      const agregarButton = document.getElementById("agregarPasajero");
      if (agregarButton) {
        agregarButton.onclick = function() {
          try {
            agregarNuevoPasajero();
          } catch (error) {
            mostrarMensaje("Error al agregar pasajero: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de agregar hotel
      const agregarHotelButton = document.getElementById("agregarHotel");
      if (agregarHotelButton) {
        agregarHotelButton.onclick = function() {
          try {
            mostrarSeccionHotel();
          } catch (error) {
            mostrarMensaje("Error al agregar hotel: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de agregar servicio
      const agregarServicioButton = document.getElementById("agregarServicio");
      if (agregarServicioButton) {
        agregarServicioButton.onclick = function() {
          try {
            agregarNuevoServicio();
          } catch (error) {
            mostrarMensaje("Error al agregar servicio: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de agregar vuelo
      const agregarVueloButton = document.getElementById("agregarVuelo");
      if (agregarVueloButton) {
        agregarVueloButton.onclick = function() {
          try {
            agregarNuevoVuelo();
          } catch (error) {
            mostrarMensaje("Error al agregar vuelo: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de eliminar hotel
      const eliminarHotelButton = document.getElementById("eliminarHotel");
      if (eliminarHotelButton) {
        eliminarHotelButton.onclick = function() {
          try {
            eliminarHotel();
          } catch (error) {
            mostrarMensaje("Error al eliminar hotel: " + error.message, "error");
          }
        };
      }
      
      // Asignar evento al botón de crear reserva
      const crearReservaButton = document.getElementById("crearReserva");
      if (crearReservaButton) {
        crearReservaButton.onclick = function() {
          ejecutarCrearReserva();
        };
        // Deshabilitar el botón inicialmente
        crearReservaButton.disabled = true;
        crearReservaButton.style.opacity = "0.5";
        crearReservaButton.style.cursor = "not-allowed";
      }
      
      // Agregar event listeners para campos de reserva
      const camposReserva = ['tipoReserva', 'estadoReserva', 'fechaViaje', 'vendedor', 'cliente', 'moneda'];
      camposReserva.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.addEventListener('change', actualizarEstadoBotonCrearReserva);
          campo.addEventListener('input', actualizarEstadoBotonCrearReserva);
        }
      });
      
      // Agregar event listeners para campos de hotel
      const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
      camposHotel.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.addEventListener('change', () => {
            validarFechasHotel();
            actualizarEstadoBotonCrearReserva();
          });
          campo.addEventListener('input', () => {
            validarFechasHotel();
            actualizarEstadoBotonCrearReserva();
          });
        }
      });
      
      // Cargar datos maestros al iniciar
      cargarDatosMaestros();
    }
  } catch (error) {
    // Error silencioso
  }
});

/**
 * Valida las fechas del hotel en tiempo real
 */
function validarFechasHotel() {
  const hotelIn = document.getElementById("hotel_in");
  const hotelOut = document.getElementById("hotel_out");
  
  // Solo validar si ambos campos tienen valores
  if (!hotelIn || !hotelOut || !hotelIn.value || !hotelOut.value) {
    // Limpiar estilos de error si no hay valores
    if (hotelIn) hotelIn.style.borderColor = '';
    if (hotelOut) hotelOut.style.borderColor = '';
    return;
  }
  
  const fechaIn = new Date(hotelIn.value);
  const fechaOut = new Date(hotelOut.value);
  
  // Verificar que las fechas sean válidas
  if (isNaN(fechaIn.getTime()) || isNaN(fechaOut.getTime())) {
    return;
  }
  
  // Verificar que la fecha de salida sea posterior a la de entrada
  if (fechaOut <= fechaIn) {
    // Mostrar error visual
    hotelIn.style.borderColor = '#d32f2f';
    hotelOut.style.borderColor = '#d32f2f';
    return;
  }
  
  // Calcular la diferencia en días
  const diferenciaDias = Math.floor((fechaOut - fechaIn) / (1000 * 60 * 60 * 24));
  
  if (diferenciaDias < 1) {
    // Mostrar error visual
    hotelIn.style.borderColor = '#d32f2f';
    hotelOut.style.borderColor = '#d32f2f';
  } else {
    // Limpiar estilos de error
    hotelIn.style.borderColor = '';
    hotelOut.style.borderColor = '';
  }
}

/**
 * Cargar datos maestros desde el servidor
 */
async function cargarDatosMaestros() {
  try {
    // Usar la variable global RPA_API_URL inyectada por webpack
    const masterDataUrl = typeof RPA_API_URL !== 'undefined'
      ? RPA_API_URL + '/api/master-data'
      : 'http://localhost:3001/api/master-data';
    
    const response = await fetch(masterDataUrl);
    
    if (!response.ok) {
      console.warn('⚠️ No se pudieron cargar los datos maestros, usando valores por defecto');
      return;
    }
    
    const result = await response.json();
    
    if (result.success && result.data) {
      masterData.sellers = result.data.sellers || [];
      masterData.clients = result.data.clients || [];
      masterData.statuses = result.data.statuses || [];
      masterData.reservationTypes = result.data.reservationTypes || [];
      masterData.genders = result.data.genders || [];
      masterData.documentTypes = result.data.documentTypes || [];
      masterData.countries = result.data.countries || [];
      masterData.loaded = true;
      
      console.log('✅ Datos maestros cargados:', {
        vendedores: masterData.sellers.length,
        clientes: masterData.clients.length,
        estados: masterData.statuses.length,
        tiposReserva: masterData.reservationTypes.length,
        generos: masterData.genders.length,
        tiposDoc: masterData.documentTypes.length,
        paises: masterData.countries.length
      });
      
      // Poblar los selects de reserva
      poblarSelectReserva();
    }
  } catch (error) {
    console.error('❌ Error cargando datos maestros:', error);
  }
}

/**
 * Poblar los selects de la sección de reserva
 */
function poblarSelectReserva() {
  // Tipo de Reserva
  const tipoReservaSelect = document.getElementById('tipoReserva');
  if (tipoReservaSelect && masterData.reservationTypes.length > 0) {
    tipoReservaSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.reservationTypes.forEach(tipo => {
      const option = document.createElement('option');
      option.value = tipo.name;
      option.textContent = tipo.name;
      tipoReservaSelect.appendChild(option);
    });
    console.log(`📋 Tipo Reserva poblado con ${masterData.reservationTypes.length} opciones`);
  } else if (tipoReservaSelect) {
    // Opciones por defecto si no hay datos maestros
    tipoReservaSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="AGENCIAS [COAG]">AGENCIAS [COAG]</option>
      <option value="MAYORISTA [COMA]">MAYORISTA [COMA]</option>
      <option value="DIRECTO [CODI]">DIRECTO [CODI]</option>
      <option value="CORPORATIVA [COCO]">CORPORATIVA [COCO]</option>
    `;
    console.log(`📋 Tipo Reserva poblado con opciones por defecto`);
  }
  
  // Estado
  const estadoSelect = document.getElementById('estadoReserva');
  if (estadoSelect && masterData.statuses.length > 0) {
    estadoSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.statuses.forEach(estado => {
      const option = document.createElement('option');
      option.value = estado.name;
      option.textContent = estado.name;
      estadoSelect.appendChild(option);
    });
    console.log(`📋 Estado poblado con ${masterData.statuses.length} opciones`);
  } else if (estadoSelect) {
    // Opciones por defecto si no hay datos maestros
    estadoSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="PENDIENTE DE CONFIRMACION [PC]">PENDIENTE DE CONFIRMACION [PC]</option>
      <option value="CONFIRMADA [CO]">CONFIRMADA [CO]</option>
      <option value="CANCELADA [CA]">CANCELADA [CA]</option>
    `;
    console.log(`📋 Estado poblado con opciones por defecto`);
  }
  
  // Vendedor
  const vendedorSelect = document.getElementById('vendedor');
  if (vendedorSelect && masterData.sellers.length > 0) {
    vendedorSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.sellers.forEach(vendedor => {
      const option = document.createElement('option');
      option.value = vendedor.name;
      option.textContent = vendedor.name;
      vendedorSelect.appendChild(option);
    });
    console.log(`📋 Vendedor poblado con ${masterData.sellers.length} opciones`);
  } else if (vendedorSelect) {
    // Opciones por defecto si no hay datos maestros
    vendedorSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="TEST TEST">TEST TEST</option>
    `;
    console.log(`📋 Vendedor poblado con opciones por defecto`);
  }
  
  // Cliente
  const clienteSelect = document.getElementById('cliente');
  if (clienteSelect && masterData.clients.length > 0) {
    clienteSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.clients.forEach(cliente => {
      const option = document.createElement('option');
      option.value = cliente.name;
      option.textContent = cliente.name;
      clienteSelect.appendChild(option);
    });
    console.log(`📋 Cliente poblado con ${masterData.clients.length} opciones`);
  } else if (clienteSelect) {
    // Opciones por defecto si no hay datos maestros
    clienteSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="DESPEGAR - TEST - 1">DESPEGAR - TEST - 1</option>
    `;
    console.log(`📋 Cliente poblado con opciones por defecto`);
  }
}

/**
 * Poblar los selects de un formulario de pasajero
 * @param {number} numero - Número del pasajero
 */
function poblarSelectsPasajero(numero) {
  // Sexo
  const sexoSelect = document.getElementById(`sexo_${numero}`);
  if (sexoSelect && masterData.genders.length > 0) {
    const valorActual = sexoSelect.value;
    sexoSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.genders.forEach(genero => {
      const option = document.createElement('option');
      option.value = genero.code;
      option.textContent = genero.name;
      sexoSelect.appendChild(option);
    });
    if (valorActual) sexoSelect.value = valorActual;
    console.log(`📋 Sexo ${numero} poblado con ${masterData.genders.length} opciones`);
  } else if (sexoSelect) {
    // Si no hay datos maestros, usar opciones por defecto
    sexoSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="M">MASCULINO</option>
      <option value="F">FEMENINO</option>
    `;
    console.log(`📋 Sexo ${numero} poblado con opciones por defecto`);
  }
  
  // Tipo de Documento
  const tipoDocSelect = document.getElementById(`tipoDoc_${numero}`);
  if (tipoDocSelect && masterData.documentTypes.length > 0) {
    const valorActual = tipoDocSelect.value;
    tipoDocSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.documentTypes.forEach(tipo => {
      const option = document.createElement('option');
      option.value = tipo.code;
      option.textContent = tipo.name;
      tipoDocSelect.appendChild(option);
    });
    if (valorActual) tipoDocSelect.value = valorActual;
    console.log(`📋 Tipo Doc ${numero} poblado con ${masterData.documentTypes.length} opciones`);
  } else if (tipoDocSelect) {
    // Si no hay datos maestros, usar opciones por defecto
    tipoDocSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="DNI">DOCUMENTO NACIONAL DE IDENTIDAD</option>
      <option value="PAS">PASAPORTE</option>
      <option value="CI">CÉDULA DE IDENTIDAD</option>
      <option value="LE">LIBRETA DE ENROLAMIENTO</option>
      <option value="LC">LIBRETA CÍVICA</option>
    `;
    console.log(`📋 Tipo Doc ${numero} poblado con opciones por defecto`);
  }
  
  // Nacionalidad
  const nacionalidadSelect = document.getElementById(`nacionalidad_${numero}`);
  if (nacionalidadSelect && masterData.countries.length > 0) {
    const valorActual = nacionalidadSelect.value;
    nacionalidadSelect.innerHTML = '<option value="">Seleccione...</option>';
    masterData.countries.forEach(pais => {
      const option = document.createElement('option');
      option.value = pais.name;
      option.textContent = pais.name;
      nacionalidadSelect.appendChild(option);
    });
    if (valorActual) nacionalidadSelect.value = valorActual;
    console.log(`📋 Nacionalidad ${numero} poblada con ${masterData.countries.length} opciones`);
  } else if (nacionalidadSelect) {
    // Si no hay datos maestros, usar opciones por defecto
    nacionalidadSelect.innerHTML = `
      <option value="">Seleccione...</option>
      <option value="ARGENTINA">ARGENTINA</option>
      <option value="BRASIL">BRASIL</option>
      <option value="CHILE">CHILE</option>
      <option value="URUGUAY">URUGUAY</option>
      <option value="PARAGUAY">PARAGUAY</option>
      <option value="BOLIVIA">BOLIVIA</option>
      <option value="PERU">PERU</option>
      <option value="COLOMBIA">COLOMBIA</option>
      <option value="VENEZUELA">VENEZUELA</option>
      <option value="ECUADOR">ECUADOR</option>
      <option value="MEXICO">MEXICO</option>
      <option value="ESPAÑA">ESPAÑA</option>
      <option value="ESTADOS UNIDOS">ESTADOS UNIDOS</option>
    `;
    console.log(`📋 Nacionalidad ${numero} poblada con opciones por defecto`);
  }
}

async function run(isReExtract = false) {
  try {
    // Ocultar el botón de extraer
    const runButton = document.getElementById("run");
    runButton.style.display = "none";
    
    // Mostrar el loader
    const loader = document.getElementById("loader");
    loader.style.display = "block";
    
    // Ocultar los resultados mientras se extrae
    const resultsDiv = document.getElementById("results");
    resultsDiv.style.display = "none";

  const item = Office.context.mailbox.item;
    
    // Obtener el cuerpo del correo
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text, 
      async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const cuerpoCorreo = result.value;
          
          try {
            // Extraer imágenes del email
            let images = [];
            try {
              images = await extraerImagenesDelEmail();
              if (images.length > 0) {
                console.log(`📸 ${images.length} imagen(es) extraída(s) del email`);
              }
            } catch (imageError) {
              console.warn('⚠️ Error al extraer imágenes, continuando sin ellas:', imageError);
              // Continuar sin imágenes para no romper el flujo
            }
            
            // Llamar al servicio de extracción con IA (incluyendo imágenes si las hay)
            const extractedData = await extraerDatosConIA(cuerpoCorreo, isReExtract, images);
            
            // Ocultar loader
            loader.style.display = "none";
            
            // Mostrar resultados
            resultsDiv.style.display = "block";
            
            // Habilitar scroll cuando hay resultados
            const mainContainer = document.querySelector(".ms-welcome__main");
            if (mainContainer) {
              mainContainer.classList.remove("no-scroll");
            }
            
            // Mostrar botón de re-extracción cuando hay resultados
            const reextractButton = document.getElementById("reextract");
            if (reextractButton) {
              reextractButton.style.display = "block";
            }
            
            // Mostrar botón de agregar vuelo siempre cuando hay resultados (para poder agregar vuelos manualmente)
            const agregarVueloButton = document.getElementById("agregarVuelo");
            if (agregarVueloButton) {
              agregarVueloButton.style.display = "block";
            }
            
            if (extractedData && extractedData.passengers && extractedData.passengers.length > 0) {
              // Convertir explícitamente a booleano (puede venir como string "true"/"false" o booleano)
              const didExtractionExistRaw = extractedData.didExtractionExist;
              const didExtractionExist = didExtractionExistRaw === true || didExtractionExistRaw === 'true' || didExtractionExistRaw === 1;
              
              // Convertir doesReservationExist a booleano
              const doesReservationExistRaw = extractedData.doesReservationExist;
              const doesReservationExist = doesReservationExistRaw === true || doesReservationExistRaw === 'true' || doesReservationExistRaw === 1;
              
              // Log para debug
              console.log('🔍 didExtractionExist recibido:', didExtractionExistRaw, 'tipo:', typeof didExtractionExistRaw);
              console.log('🔍 didExtractionExist convertido:', didExtractionExist);
              console.log('🔍 doesReservationExist recibido:', doesReservationExistRaw, 'tipo:', typeof doesReservationExistRaw);
              console.log('🔍 doesReservationExist convertido:', doesReservationExist);
              
              // Guardar el estado de la extracción globalmente
              extractionState.didExtractionExist = didExtractionExist;
              extractionState.doesReservationExist = doesReservationExist;
              extractionState.reservationCode = extractedData.reservationCode || null;
              
              // Guardar la data original completa de la extracción (para comparación al editar)
              // Solo si existe una reserva (doesReservationExist), no solo una extracción
              if (doesReservationExist) {
                extractionState.originData = JSON.parse(JSON.stringify(extractedData)); // Deep copy
              }
              
              // Log para verificar que se guardó correctamente
              console.log('💾 extractionState.didExtractionExist guardado:', extractionState.didExtractionExist);
              console.log('💾 extractionState.doesReservationExist guardado:', extractionState.doesReservationExist);
              console.log('💾 extractionState.reservationCode guardado:', extractionState.reservationCode);
              console.log('💾 extractionState.originData guardado:', extractionState.originData);
              
              // Crear formularios según el número de pasajeros extraídos
              crearFormulariosPasajeros(extractedData.passengers.length, didExtractionExist);
              
              // Llenar los datos de los pasajeros
              llenarDatosPasajeros(extractedData.passengers);
              
              // Llenar los datos de la reserva
              llenarDatosReserva(extractedData);
              
              // Actualizar el texto del botón según si es crear o editar (usar doesReservationExist)
              actualizarTextoBotonReserva(doesReservationExist);
              
              if(didExtractionExist) {
                mostrarMensaje(`✅ Hay una extracción de datos existente para el correo`, "success");
              }else{
                mostrarMensaje(`✅ Datos extraídos: ${extractedData.passengers.length} pasajero(s)`, "success");
              }
            } else {
              // Si no se extrajeron pasajeros, crear un formulario vacío
              crearFormulariosPasajeros(1);
              // Mostrar botón de agregar vuelo siempre cuando hay resultados (para poder agregar vuelos manualmente)
              const agregarVueloButton = document.getElementById("agregarVuelo");
              if (agregarVueloButton) {
                agregarVueloButton.style.display = "block";
              }
              mostrarMensaje("No se pudieron extraer datos. Por favor, llena el formulario manualmente.", "info");
            }
          } catch (error) {
            // Ocultar loader
            loader.style.display = "none";
            
            // Mostrar resultados con formulario vacío
            resultsDiv.style.display = "block";
            
            // Habilitar scroll cuando hay resultados (incluso si es error)
            const mainContainer = document.querySelector(".ms-welcome__main");
            if (mainContainer) {
              mainContainer.classList.remove("no-scroll");
            }
            
            // Mostrar botón de agregar vuelo siempre cuando hay resultados (para poder agregar vuelos manualmente)
            const agregarVueloButton = document.getElementById("agregarVuelo");
            if (agregarVueloButton) {
              agregarVueloButton.style.display = "block";
            }
            
            // Si falla la extracción, crear un formulario vacío
            crearFormulariosPasajeros(1);
            mostrarMensaje("Error al extraer datos: " + error.message + ". Llena el formulario manualmente.", "error");
          }
        } else {
          // Ocultar loader
          loader.style.display = "none";
          
          // Mostrar botón de nuevo
          runButton.style.display = "block";
          
          mostrarMensaje("Error al obtener el contenido del correo", "error");
        }
      }
    );
  } catch (error) {
    // Ocultar loader
    const loader = document.getElementById("loader");
    loader.style.display = "none";
    
    // Mostrar botón de nuevo
    const runButton = document.getElementById("run");
    runButton.style.display = "block";
    
    mostrarMensaje("Error inesperado: " + error.message, "error");
  }
}

/**
 * Extrae imágenes del email actual (attachments e inline)
 * @returns {Promise<File[]>} Array de archivos de imagen
 */
async function extraerImagenesDelEmail() {
  const images = [];
  const item = Office.context.mailbox.item;
  const MAX_IMAGE_SIZE = 10 * 1024 * 1024; // 10MB por imagen
  const MAX_TOTAL_SIZE = 50 * 1024 * 1024; // 50MB total
  let totalSize = 0;
  
  try {
    // Validar que el item existe
    if (!item) {
      console.warn('⚠️ No se pudo acceder al item del email');
      return [];
    }
    // 1. Extraer attachments de tipo imagen
    if (item.attachments && item.attachments.length > 0) {
      for (let i = 0; i < item.attachments.length; i++) {
        const attachment = item.attachments[i];
        
        // Verificar si es una imagen por extensión
        const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg'];
        const isImage = imageExtensions.some(ext => 
          attachment.name.toLowerCase().endsWith(ext)
        );
        
        if (isImage) {
          try {
            // Obtener el attachment como base64
            const attachmentData = await new Promise((resolve, reject) => {
              item.getAttachmentContentAsync(attachment.id, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  resolve(result.value);
                } else {
                  reject(new Error(result.error.message));
                }
              });
            });
            
            // Convertir base64 a Blob
            const contentType = attachmentData.contentType || 'image/jpeg';
            const base64Data = attachmentData.content;
            const binaryString = atob(base64Data);
            const bytes = new Uint8Array(binaryString.length);
            for (let j = 0; j < binaryString.length; j++) {
              bytes[j] = binaryString.charCodeAt(j);
            }
            const blob = new Blob([bytes], { type: contentType });
            
            // Validar tamaño individual
            if (blob.size > MAX_IMAGE_SIZE) {
              console.warn(`⚠️ Imagen ${attachment.name} es muy grande (${(blob.size / 1024 / 1024).toFixed(2)}MB), se comprimirá`);
              const compressedBlob = await comprimirImagen(blob, attachment.name);
              if (compressedBlob && compressedBlob.size > 0) {
                // Validar tamaño total
                if (totalSize + compressedBlob.size > MAX_TOTAL_SIZE) {
                  console.warn(`⚠️ Límite total de tamaño alcanzado. Se omitirá ${attachment.name}`);
                  continue;
                }
                const file = new File([compressedBlob], attachment.name, { type: compressedBlob.type });
                images.push(file);
                totalSize += compressedBlob.size;
              } else {
                console.warn(`⚠️ No se pudo comprimir ${attachment.name}, se omitirá`);
              }
            } else {
              // Validar tamaño total
              if (totalSize + blob.size > MAX_TOTAL_SIZE) {
                console.warn(`⚠️ Límite total de tamaño alcanzado. Se omitirá ${attachment.name}`);
                continue;
              }
              const file = new File([blob], attachment.name, { type: contentType });
              images.push(file);
              totalSize += blob.size;
            }
          } catch (error) {
            console.warn(`⚠️ Error al extraer attachment ${attachment.name}:`, error);
            // Continuar con otros attachments sin romper el flujo
          }
        }
      }
    }
    
    // 2. Extraer imágenes inline del HTML del body
    try {
      const htmlBody = await new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Html, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(new Error(result.error.message));
          }
        });
      });
      
      // Parsear HTML y buscar imágenes
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlBody, 'text/html');
      const imgTags = doc.querySelectorAll('img');
      
      for (const img of imgTags) {
        const src = img.getAttribute('src');
        if (!src) continue;
        
        // Si es base64 (data:image/...)
        if (src.startsWith('data:image/')) {
          try {
            const [header, base64Data] = src.split(',');
            const contentTypeMatch = header.match(/data:image\/([^;]+)/);
            const contentType = contentTypeMatch ? `image/${contentTypeMatch[1]}` : 'image/png';
            
            const binaryString = atob(base64Data);
            const bytes = new Uint8Array(binaryString.length);
            for (let j = 0; j < binaryString.length; j++) {
              bytes[j] = binaryString.charCodeAt(j);
            }
            const blob = new Blob([bytes], { type: contentType });
            
            // Validar tamaño individual
            if (blob.size > MAX_IMAGE_SIZE) {
              console.warn(`⚠️ Imagen inline es muy grande (${(blob.size / 1024 / 1024).toFixed(2)}MB), se comprimirá`);
              const compressedBlob = await comprimirImagen(blob, `inline-image-${images.length}.png`);
              if (compressedBlob && compressedBlob.size > 0) {
                // Validar tamaño total
                if (totalSize + compressedBlob.size > MAX_TOTAL_SIZE) {
                  console.warn(`⚠️ Límite total de tamaño alcanzado. Se omitirá imagen inline`);
                  continue;
                }
                const file = new File([compressedBlob], `inline-image-${images.length}.png`, { type: compressedBlob.type });
                images.push(file);
                totalSize += compressedBlob.size;
              } else {
                console.warn(`⚠️ No se pudo comprimir imagen inline, se omitirá`);
              }
            } else {
              // Validar tamaño total
              if (totalSize + blob.size > MAX_TOTAL_SIZE) {
                console.warn(`⚠️ Límite total de tamaño alcanzado. Se omitirá imagen inline`);
                continue;
              }
              const fileName = `inline-image-${images.length}.${contentType.split('/')[1] || 'png'}`;
              const file = new File([blob], fileName, { type: contentType });
              images.push(file);
              totalSize += blob.size;
            }
          } catch (error) {
            console.warn(`⚠️ Error al procesar imagen inline:`, error);
            // Continuar sin romper el flujo
          }
        }
        // Si es una URL externa, no la podemos extraer directamente
        // (requeriría permisos adicionales o Graph API)
      }
    } catch (error) {
      console.warn('⚠️ Error al extraer imágenes inline del HTML:', error);
      // Continuar sin romper el flujo
    }
    
    console.log(`📸 Imágenes extraídas: ${images.length} (${(totalSize / 1024 / 1024).toFixed(2)}MB total)`);
    
    if (images.length > 0 && totalSize > MAX_TOTAL_SIZE * 0.8) {
      console.warn(`⚠️ Advertencia: Se está usando ${((totalSize / MAX_TOTAL_SIZE) * 100).toFixed(1)}% del límite total de tamaño`);
    }
    
    return images;
    
  } catch (error) {
    console.error('❌ Error al extraer imágenes:', error);
    // Retornar array vacío para no romper el flujo principal
    // No mostrar error al usuario ya que las imágenes son opcionales
    return [];
  }
}

/**
 * Comprime una imagen usando Canvas API
 * @param {Blob} imageBlob - Blob de la imagen
 * @param {string} fileName - Nombre del archivo
 * @returns {Promise<Blob|null>} Blob comprimido o null si falla
 */
async function comprimirImagen(imageBlob, fileName) {
  return new Promise((resolve) => {
    try {
      const img = new Image();
      const url = URL.createObjectURL(imageBlob);
      
      // Timeout para evitar que se quede colgado
      const timeout = setTimeout(() => {
        URL.revokeObjectURL(url);
        console.warn('⚠️ Timeout al comprimir imagen');
        resolve(null);
      }, 30000); // 30 segundos
      
      img.onload = () => {
        clearTimeout(timeout);
        URL.revokeObjectURL(url);
        
        try {
          // Calcular nuevo tamaño (máximo 1920px de ancho o alto)
          const maxDimension = 1920;
          let width = img.width;
          let height = img.height;
          
          if (width > maxDimension || height > maxDimension) {
            if (width > height) {
              height = (height / width) * maxDimension;
              width = maxDimension;
            } else {
              width = (width / height) * maxDimension;
              height = maxDimension;
            }
          }
          
          // Validar dimensiones mínimas
          if (width < 1 || height < 1) {
            console.warn('⚠️ Dimensiones inválidas para compresión');
            resolve(null);
            return;
          }
          
          // Crear canvas y comprimir
          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          
          if (!ctx) {
            console.warn('⚠️ No se pudo obtener contexto del canvas');
            resolve(null);
            return;
          }
          
          ctx.drawImage(img, 0, 0, width, height);
          
          // Convertir a blob con calidad 0.8
          canvas.toBlob((blob) => {
            if (blob && blob.size > 0) {
              resolve(blob);
            } else {
              console.warn('⚠️ Error al generar blob comprimido');
              resolve(null);
            }
          }, 'image/jpeg', 0.8);
        } catch (error) {
          console.warn('⚠️ Error durante la compresión:', error);
          resolve(null);
        }
      };
      
      img.onerror = () => {
        clearTimeout(timeout);
        URL.revokeObjectURL(url);
        console.warn('⚠️ Error al cargar imagen para compresión');
        resolve(null);
      };
      
      img.src = url;
    } catch (error) {
      console.warn('⚠️ Error al iniciar compresión:', error);
      resolve(null);
    }
  });
}

/**
 * Llama al servicio de extracción con IA
 * @param {string} emailContent - Contenido del email
 * @param {boolean} isReExtract - Si es re-extracción
 * @param {File[]} images - Array de imágenes extraídas (opcional)
 */
async function extraerDatosConIA(emailContent, isReExtract = false, images = []) {
  // Usar la variable global RPA_API_URL inyectada por webpack
  const extractUrl = typeof RPA_API_URL !== 'undefined'
    ? RPA_API_URL + '/api/extract'
    : 'http://localhost:3001/api/extract';
  const mailbox = Office.context.mailbox;
  
  // Log para verificar que isReExtract se está enviando
  console.log('📤 Enviando extracción - isReExtract:', isReExtract, 'tipo:', typeof isReExtract);
  console.log('📸 Imágenes a enviar:', images.length);
  
  // Si hay imágenes, usar FormData; si no, usar JSON
  if (images && images.length > 0) {
    try {
      const formData = new FormData();
      formData.append('emailContent', emailContent);
      formData.append('userId', mailbox.userProfile.emailAddress || 'outlook-user');
      formData.append('conversationId', mailbox.item.conversationId || 'conversation-id');
      formData.append('isReExtract', isReExtract.toString());
      
      // Validar y agregar cada imagen
      let validImagesCount = 0;
      images.forEach((image, index) => {
        // Validar que la imagen existe y tiene tamaño válido
        if (image && image.size > 0 && image.size <= 50 * 1024 * 1024) {
          formData.append('images', image, image.name);
          validImagesCount++;
        } else {
          console.warn(`⚠️ Imagen ${image?.name || index} inválida o muy grande, se omitirá`);
        }
      });
      
      if (validImagesCount === 0) {
        console.warn('⚠️ No hay imágenes válidas, se enviará sin imágenes');
        // Continuar con JSON si no hay imágenes válidas
        images = [];
      } else {
        console.log(`📤 Enviando ${validImagesCount} imagen(es) con FormData`);
        
        const response = await fetch(extractUrl, {
          method: 'POST',
          body: formData
          // No establecer Content-Type, el navegador lo hará automáticamente con el boundary
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({ error: 'Error al extraer datos' }));
          throw new Error(errorData.error || 'Error al extraer datos');
        }

        const result = await response.json();
        
        // Incluir didExtractionExist y doesReservationExist del nivel superior en el objeto data
        const dataWithExtractionState = {
          ...result.data,
          didExtractionExist: result.didExtractionExist || false,
          doesReservationExist: result.doesReservationExist || false
        };
        
        return dataWithExtractionState;
      }
    } catch (error) {
      console.error('❌ Error al enviar FormData con imágenes:', error);
      // Si falla el envío con FormData, intentar sin imágenes
      console.warn('⚠️ Reintentando sin imágenes...');
      images = [];
    }
  }
  
  // Sin imágenes o fallback, usar JSON como antes
  const response = await fetch(extractUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      emailContent: emailContent,
      userId: mailbox.userProfile.emailAddress || 'outlook-user',
      conversationId: mailbox.item.conversationId || 'conversation-id',
      isReExtract: isReExtract
    })
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error || 'Error al extraer datos');
  }

  const result = await response.json();
  
  // Incluir didExtractionExist y doesReservationExist del nivel superior en el objeto data
  const dataWithExtractionState = {
    ...result.data,
    didExtractionExist: result.didExtractionExist || false,
    doesReservationExist: result.doesReservationExist || false
  };
  
  return dataWithExtractionState;
}

function crearFormulariosPasajeros(numeroPasajeros, didExtractionExist = false) {
  const container = document.getElementById("pasajerosContainer");
  container.innerHTML = ""; // Limpiar contenedor
  
  for (let i = 0; i < numeroPasajeros; i++) {
    const pasajeroDiv = crearFormularioPasajero(i + 1);
    container.appendChild(pasajeroDiv);
  }
  
  // Actualizar estado del botón después de crear los formularios
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

function crearFormularioPasajero(numero) {
  const pasajeroDiv = document.createElement("div");
  pasajeroDiv.className = "pasajero-acordeon";
  pasajeroDiv.dataset.numeroPasajero = numero;
  
  // Cabecera del acordeón (clickeable)
  const header = document.createElement("div");
  header.className = "pasajero-header";
  header.innerHTML = `
    <span class="pasajero-titulo">Pasajero ${numero}</span>
    <div class="pasajero-actions">
      <span class="arrow">▼</span>
      <button class="btn-eliminar-pasajero" title="Eliminar pasajero">✕</button>
    </div>
  `;
  
  // Contenido del acordeón (formulario)
  const content = document.createElement("div");
  content.className = "pasajero-content";
  content.style.display = "none";
  content.innerHTML = `
    <div>
      <label>Tipo de Pasajero:</label>
      <select id="tipoPasajero_${numero}">
        <option value="">Seleccione...</option>
        <option value="adulto">Adulto</option>
        <option value="menor">Menor</option>
        <option value="infante">Infante</option>
      </select>
    </div>

    <div>
      <label>Nombre: <span style="color: red;">*</span></label>
      <input type="text" id="nombre_${numero}" placeholder="Ingrese el nombre">
    </div>

    <div>
      <label>Apellido: <span style="color: red;">*</span></label>
      <input type="text" id="apellido_${numero}" placeholder="Ingrese el apellido">
    </div>

    <div>
      <label>DNI:</label>
      <input type="number" id="dni_${numero}" placeholder="Ingrese el DNI">
    </div>

    <div>
      <label>Fecha de Nacimiento:</label>
      <input type="date" id="fechaNacimiento_${numero}">
    </div>

    <div>
      <label>CUIL:</label>
      <input type="number" id="cuil_${numero}" placeholder="Ingrese el CUIL">
    </div>

    <div>
      <label>Tipo de Documento:</label>
      <select id="tipoDoc_${numero}">
        <option value="">Seleccione...</option>
        <option value="dni">DNI</option>
        <option value="pasaporte">Pasaporte</option>
        <option value="cedula">Cédula</option>
        <option value="otro">Otro</option>
      </select>
    </div>

    <div>
      <label>Sexo:</label>
      <select id="sexo_${numero}">
        <option value="">Seleccione...</option>
        <option value="masculino">Masculino</option>
        <option value="femenino">Femenino</option>
        <option value="otro">Otro</option>
      </select>
    </div>

    <div>
      <label>Nacionalidad:</label>
      <select id="nacionalidad_${numero}">
        <option value="">Seleccione...</option>
        <option value="argentina">Argentina</option>
        <option value="brasilera">Brasilera</option>
        <option value="chilena">Chilena</option>
        <option value="uruguaya">Uruguaya</option>
        <option value="paraguaya">Paraguaya</option>
        <option value="boliviana">Boliviana</option>
        <option value="peruana">Peruana</option>
        <option value="colombiana">Colombiana</option>
        <option value="venezolana">Venezolana</option>
        <option value="otra">Otra</option>
      </select>
    </div>

    <div>
      <label>Dirección:</label>
      <input type="text" id="direccion_${numero}" placeholder="Ingrese la dirección">
    </div>

    <div>
      <label>Número de Teléfono:</label>
      <input type="tel" id="telefono_${numero}" placeholder="Ingrese el teléfono">
    </div>
  `;
  
  // Funcionalidad de acordeón (toggle)
  header.onclick = function(e) {
    // No hacer toggle si se clickeó el botón de eliminar
    if (e.target.classList.contains('btn-eliminar-pasajero')) {
      return;
    }
    
    const isOpen = content.style.display === "block";
    const arrow = header.querySelector(".arrow");
    
    if (isOpen) {
      content.style.display = "none";
      arrow.style.transform = "rotate(0deg)";
    } else {
      content.style.display = "block";
      arrow.style.transform = "rotate(180deg)";
    }
  };
  
  // Agregar event listeners para validación en tiempo real (solo nombre y apellido)
  setTimeout(() => {
    const nombreInput = document.getElementById(`nombre_${numero}`);
    const apellidoInput = document.getElementById(`apellido_${numero}`);
    
    if (nombreInput) {
      nombreInput.addEventListener('input', actualizarEstadoBotonCrearReserva);
    }
    if (apellidoInput) {
      apellidoInput.addEventListener('input', actualizarEstadoBotonCrearReserva);
    }
  }, 100);
  
  // Funcionalidad del botón eliminar
  const btnEliminar = header.querySelector('.btn-eliminar-pasajero');
  btnEliminar.onclick = function(e) {
    e.stopPropagation(); // Evitar que se abra/cierre el acordeón
    eliminarPasajero(pasajeroDiv);
  };
  
  pasajeroDiv.appendChild(header);
  pasajeroDiv.appendChild(content);
  
  // Poblar los selects con datos maestros después de crear el formulario
  setTimeout(() => {
    poblarSelectsPasajero(numero);
  }, 50);
  
  return pasajeroDiv;
}

function eliminarPasajero(pasajeroDiv) {
  try {
    const container = document.getElementById("pasajerosContainer");
    const pasajeros = container.querySelectorAll(".pasajero-acordeon");
    
    // No permitir eliminar si solo hay un pasajero
    if (pasajeros.length <= 1) {
      mostrarMensaje("Debe haber al menos un pasajero", "info");
      return;
    }
    
    // Eliminar directamente sin confirmación (o puedes crear un modal personalizado)
    pasajeroDiv.remove();
    renumerarPasajeros();
    mostrarMensaje("Pasajero eliminado correctamente", "success");
    
    // Actualizar estado del botón después de eliminar
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  } catch (error) {
    mostrarMensaje("Error al eliminar pasajero", "error");
  }
}

function renumerarPasajeros() {
  const container = document.getElementById("pasajerosContainer");
  const pasajeros = container.querySelectorAll(".pasajero-acordeon");
  
  pasajeros.forEach((pasajeroDiv, index) => {
    const nuevoNumero = index + 1;
    const titulo = pasajeroDiv.querySelector(".pasajero-titulo");
    titulo.textContent = `Pasajero ${nuevoNumero}`;
    pasajeroDiv.dataset.numeroPasajero = nuevoNumero;
  });
}

function agregarNuevoPasajero() {
  try {
    const container = document.getElementById("pasajerosContainer");
    if (!container) {
      return;
    }
    
    const pasajeros = container.querySelectorAll(".pasajero-acordeon");
    const nuevoNumero = pasajeros.length + 1;
    
    const nuevoPasajero = crearFormularioPasajero(nuevoNumero);
    container.appendChild(nuevoPasajero);
    
    // Abrir automáticamente el nuevo pasajero
    const content = nuevoPasajero.querySelector(".pasajero-content");
    const arrow = nuevoPasajero.querySelector(".arrow");
    if (content && arrow) {
      content.style.display = "block";
      arrow.style.transform = "rotate(180deg)";
    }
    
    // Scroll suave hacia el nuevo pasajero
    if (nuevoPasajero.scrollIntoView) {
      nuevoPasajero.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    // Actualizar estado del botón después de agregar
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  } catch (error) {
    throw error;
  }
}

async function guardarDatos() {
  try {
    const container = document.getElementById("pasajerosContainer");
    if (!container) {
      mostrarMensaje("No se encontró el contenedor de pasajeros", "error");
      return;
    }
    
    const pasajeros = container.querySelectorAll(".pasajero-acordeon");
    if (pasajeros.length === 0) {
      mostrarMensaje("No hay pasajeros para guardar. Por favor, extraiga datos primero.", "info");
      return;
    }
    
    // Recopilar datos de todos los pasajeros
    const todosPasajeros = [];
    pasajeros.forEach((pasajeroDiv, index) => {
      const numero = pasajeroDiv.dataset.numeroPasajero;
      const content = pasajeroDiv.querySelector(".pasajero-content");
      
      if (content) {
        const datos = {
          numeroPasajero: index + 1,
          tipoPasajero: content.querySelector(`#tipoPasajero_${numero}`)?.value || "",
          nombre: content.querySelector(`#nombre_${numero}`)?.value || "",
          apellido: content.querySelector(`#apellido_${numero}`)?.value || "",
          dni: content.querySelector(`#dni_${numero}`)?.value || "",
          fechaNacimiento: content.querySelector(`#fechaNacimiento_${numero}`)?.value || "",
          cuil: content.querySelector(`#cuil_${numero}`)?.value || "",
          tipoDoc: content.querySelector(`#tipoDoc_${numero}`)?.value || "",
          sexo: content.querySelector(`#sexo_${numero}`)?.value || "",
          nacionalidad: content.querySelector(`#nacionalidad_${numero}`)?.value || "",
          direccion: content.querySelector(`#direccion_${numero}`)?.value || "",
          telefono: content.querySelector(`#telefono_${numero}`)?.value || ""
        };
        
        todosPasajeros.push(datos);
      }
    });
    
    // Capturar datos de la reserva
    const datosReserva = {
      tipoReserva: document.getElementById("tipoReserva")?.value || "",
      estadoReserva: document.getElementById("estadoReserva")?.value || "",
      fechaViaje: document.getElementById("fechaViaje")?.value || "",
      vendedor: document.getElementById("vendedor")?.value || "",
      cliente: document.getElementById("cliente")?.value || "",
      codigo: document.getElementById("codigo")?.value || "",
      reservationDate: document.getElementById("fechaReserva")?.value || "",
      tourEndDate: document.getElementById("fechaFinTour")?.value || "",
      dueDate: document.getElementById("fechaVencimiento")?.value || "",
      contact: document.getElementById("contacto")?.value || "",
      contactEmail: document.getElementById("contactEmail")?.value || "",
      contactPhone: document.getElementById("contactPhone")?.value || "",
      currency: document.getElementById("moneda")?.value || "",
      exchangeRate: parseFloat(document.getElementById("tipoCambio")?.value) || 0,
      commission: parseFloat(document.getElementById("comision")?.value) || 0,
      netAmount: parseFloat(document.getElementById("montoNeto")?.value) || 0,
      grossAmount: parseFloat(document.getElementById("montoBruto")?.value) || 0,
      tripName: document.getElementById("nombreViaje")?.value || "",
      productCode: document.getElementById("codigoProducto")?.value || "",
      adults: parseInt(document.getElementById("adultos")?.value) || 0,
      children: parseInt(document.getElementById("ninos")?.value) || 0,
      infants: parseInt(document.getElementById("infantes")?.value) || 0,
      provider: document.getElementById("proveedor")?.value || "",
      reservationCode: document.getElementById("codigoReserva")?.value || extractionState.reservationCode || "",
      hotel: (() => {
        // Verificar primero si la sección está visible
        const hotelSection = document.getElementById("hotelSection");
        if (!hotelSection || hotelSection.style.display === "none") {
          return null;
        }
        
        const nombreHotel = document.getElementById("hotel_nombre")?.value || "";
        const tipoHabitacion = document.getElementById("hotel_tipo_habitacion")?.value || "";
        const ciudad = document.getElementById("hotel_ciudad")?.value || "";
        const categoria = document.getElementById("hotel_categoria")?.value || null;
        const prioridad = document.getElementById("hotel_prioridad")?.value || null;
        const hotelIn = document.getElementById("hotel_in")?.value || "";
        const hotelOut = document.getElementById("hotel_out")?.value || "";
        
        if (!nombreHotel && !tipoHabitacion && !ciudad && !hotelIn && !hotelOut) {
          return null;
        }
        
        return {
          nombre_hotel: nombreHotel,
          tipo_habitacion: tipoHabitacion,
          Ciudad: ciudad,
          Categoria: categoria || null,
          Prioridad: prioridad || null,
          in: hotelIn,
          out: hotelOut
        };
      })(),
      checkIn: document.getElementById("checkIn")?.value || "",
      checkOut: document.getElementById("checkOut")?.value || "",
      estadoDeuda: document.getElementById("estadoDeuda")?.value || ""
    };
    
    // Agregar conversationId e itemId del email actual
    const item = Office.context.mailbox.item;
    datosReserva.conversationId = item.conversationId || null;
    datosReserva.itemId = item.itemId || null;
    
    // Capturar servicios
    const servicios = [];
    const serviciosSection = document.getElementById("serviciosSection");
    const serviciosContainer = document.getElementById("serviciosContainer");
    // Solo capturar servicios si la sección está visible y hay elementos
    if (serviciosSection && serviciosSection.style.display !== "none" && serviciosContainer) {
      const servicioItems = serviciosContainer.querySelectorAll(".servicio-item");
      if (servicioItems.length > 0) {
        servicioItems.forEach((item, index) => {
          servicios.push({
            destino: document.getElementById(`servicio_destino_${index}`)?.value || "",
            in: document.getElementById(`servicio_in_${index}`)?.value || "",
            out: document.getElementById(`servicio_out_${index}`)?.value || "",
            nts: parseInt(document.getElementById(`servicio_nts_${index}`)?.value) || 0,
            basePax: parseInt(document.getElementById(`servicio_basePax_${index}`)?.value) || 0,
            servicio: document.getElementById(`servicio_servicio_${index}`)?.value || "",
            descripcion: document.getElementById(`servicio_descripcion_${index}`)?.value || "",
            estado: document.getElementById(`servicio_estado_${index}`)?.value || "",
            categoria: document.getElementById(`servicio_categoria_${index}`)?.value || "",
            prioridad: document.getElementById(`servicio_prioridad_${index}`)?.value || ""
          });
        });
      }
    }
    // Siempre asignar array vacío si no hay servicios
    datosReserva.services = servicios;
    
    // Capturar vuelos
    const vuelos = [];
    const vuelosSection = document.getElementById("vuelosSection");
    const vuelosContainer = document.getElementById("vuelosContainer");
    // Solo capturar vuelos si la sección está visible y hay elementos
    if (vuelosSection && vuelosSection.style.display !== "none" && vuelosContainer) {
      const vueloItems = vuelosContainer.querySelectorAll(".vuelo-item");
      if (vueloItems.length > 0) {
        vueloItems.forEach((item, index) => {
          vuelos.push({
            airline: document.getElementById(`vuelo_airline_${index}`)?.value || "",
            flightNumber: document.getElementById(`vuelo_flightNumber_${index}`)?.value || "",
            origin: document.getElementById(`vuelo_origin_${index}`)?.value || "",
            destination: document.getElementById(`vuelo_destination_${index}`)?.value || "",
            departureDate: document.getElementById(`vuelo_departureDate_${index}`)?.value || "",
            departureTime: document.getElementById(`vuelo_departureTime_${index}`)?.value || "",
            arrivalDate: document.getElementById(`vuelo_arrivalDate_${index}`)?.value || "",
            arrivalTime: document.getElementById(`vuelo_arrivalTime_${index}`)?.value || ""
          });
        });
      }
    }
    // Siempre asignar array vacío si no hay vuelos
    datosReserva.flights = vuelos;
    
    // Preparar datos para enviar
    const datosParaEnviar = {
      passengers: todosPasajeros,
      ...datosReserva
    };
    
    // Mostrar mensaje de procesamiento
    mostrarMensaje("Guardando datos... Por favor espere.", "info");
    
    // Usar la variable global RPA_API_URL inyectada por webpack
    const updateUrl = typeof RPA_API_URL !== 'undefined'
      ? RPA_API_URL + '/api/extract/update'
      : 'http://localhost:3001/api/extract/update';
    
    // Enviar datos al backend
    const response = await fetch(updateUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(datosParaEnviar)
    });
    
    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ error: 'Error al guardar datos' }));
      throw new Error(errorData.error || 'Error al guardar datos');
    }
    
    const result = await response.json();
    
    mostrarMensaje(`Datos de ${todosPasajeros.length} pasajero(s) guardados correctamente`, "success");
  } catch (error) {
    mostrarMensaje("Error al guardar datos: " + error.message, "error");
    throw error;
  }
}

function llenarDatosPasajeros(datosPasajeros) {
  // Función auxiliar para llenar los datos extraídos por la IA
  // Esperar un poco para asegurar que los selects estén poblados
  setTimeout(() => {
    datosPasajeros.forEach((datos, index) => {
      const numero = index + 1;
      
      if (document.getElementById(`tipoPasajero_${numero}`)) {
        // Mapear passengerType o paxType de la IA a tipoPasajero del formulario
        let tipoPasajero = "";
        const passengerType = datos.passengerType || datos.paxType;
        if (passengerType === "ADU") tipoPasajero = "adulto";
        else if (passengerType === "CHD") tipoPasajero = "menor";
        else if (passengerType === "INF") tipoPasajero = "infante";
        
        // Los valores de sex, documentType y nationality ahora vienen normalizados del backend
        // sex viene como código: "M" o "F"
        // documentType viene como código: "DNI", "PAS", "CI", etc.
        // nationality viene como nombre completo en mayúsculas: "ARGENTINA", "BRASIL", etc.
        
        document.getElementById(`tipoPasajero_${numero}`).value = tipoPasajero;
        document.getElementById(`nombre_${numero}`).value = datos.firstName || "";
        document.getElementById(`apellido_${numero}`).value = datos.lastName || "";
        document.getElementById(`dni_${numero}`).value = datos.documentNumber || "";
        // El backend envía dateOfBirth, no birthDate
        document.getElementById(`fechaNacimiento_${numero}`).value = datos.dateOfBirth || datos.birthDate || "";
        document.getElementById(`cuil_${numero}`).value = datos.cuilCuit || "";
        
        // Usar los códigos directamente tal como vienen del backend
        const tipoDocSelect = document.getElementById(`tipoDoc_${numero}`);
        if (tipoDocSelect && datos.documentType) {
          tipoDocSelect.value = datos.documentType;
          console.log(`✅ Tipo Doc ${numero}: ${datos.documentType} -> ${tipoDocSelect.value}`);
        }
        
        const sexoSelect = document.getElementById(`sexo_${numero}`);
        if (sexoSelect && datos.sex) {
          sexoSelect.value = datos.sex;
          console.log(`✅ Sexo ${numero}: ${datos.sex} -> ${sexoSelect.value}`);
        }
        
        const nacionalidadSelect = document.getElementById(`nacionalidad_${numero}`);
        if (nacionalidadSelect && datos.nationality) {
          nacionalidadSelect.value = datos.nationality;
          console.log(`✅ Nacionalidad ${numero}: ${datos.nationality} -> ${nacionalidadSelect.value}`);
        }
        
        document.getElementById(`direccion_${numero}`).value = datos.direccion || "";
        // El backend envía phoneNumber, no telefono
        document.getElementById(`telefono_${numero}`).value = datos.phoneNumber || datos.telefono || "";
      }
    });
    
    // Actualizar estado del botón después de llenar los datos
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  }, 150); // Esperar 150ms para que los selects se pueblen primero
}

/**
 * Llena los datos de la reserva extraídos por la IA
 */
function llenarDatosReserva(datosExtraidos) {
  // Repoblar los selects de reserva primero para asegurar que tengan opciones
  poblarSelectReserva();
  
  // Esperar un poco más para asegurar que los selects estén poblados
  setTimeout(() => {
    // Tipo de Reserva
    const tipoReservaSelect = document.getElementById("tipoReserva");
    if (tipoReservaSelect) {
      const valorTipoReserva = datosExtraidos.reservationType;
      console.log(`🔍 Intentando asignar Tipo Reserva: "${valorTipoReserva}"`);
      console.log(`📋 Opciones en select:`, Array.from(tipoReservaSelect.options).map(o => `"${o.value}"`));
      
      if (valorTipoReserva && valorTipoReserva !== 'null' && valorTipoReserva !== null) {
        tipoReservaSelect.value = valorTipoReserva;
        console.log(`✅ Tipo Reserva asignado: "${valorTipoReserva}" -> "${tipoReservaSelect.value}"`);
        
        // Si no se seleccionó, intentar buscar una coincidencia parcial
        if (!tipoReservaSelect.value || tipoReservaSelect.value === "") {
          const opciones = Array.from(tipoReservaSelect.options);
          const coincidencia = opciones.find(opt => 
            opt.value && valorTipoReserva &&
            (opt.value.toUpperCase().includes(valorTipoReserva.toUpperCase()) ||
            valorTipoReserva.toUpperCase().includes(opt.value.toUpperCase()))
          );
          if (coincidencia) {
            tipoReservaSelect.value = coincidencia.value;
            console.log(`✅ Tipo Reserva (coincidencia): "${valorTipoReserva}" -> "${tipoReservaSelect.value}"`);
          } else {
            console.warn(`⚠️ No se encontró coincidencia para Tipo Reserva: "${valorTipoReserva}"`);
          }
        }
      }
    }
    
    // Estado
    const estadoReservaSelect = document.getElementById("estadoReserva");
    if (estadoReservaSelect) {
      const valorEstado = datosExtraidos.status;
      const opcionesDisponibles = Array.from(estadoReservaSelect.options).map(o => o.value);
      
      console.log(`🔍 Intentando asignar Estado: "${valorEstado}"`);
      console.log(`📋 Opciones disponibles en select Estado:`, opcionesDisponibles);
      console.log(`📊 Total de opciones: ${opcionesDisponibles.length}`);
      
      if (valorEstado && valorEstado !== 'null' && valorEstado !== null) {
        // Intentar asignación directa
        estadoReservaSelect.value = valorEstado;
        console.log(`🔄 Intento directo: "${valorEstado}" -> "${estadoReservaSelect.value}"`);
        
        // Si no se seleccionó, intentar buscar una coincidencia
        if (!estadoReservaSelect.value || estadoReservaSelect.value === "") {
          console.log(`⚠️ Asignación directa falló, buscando coincidencia...`);
          
          const opciones = Array.from(estadoReservaSelect.options);
          
          // Intentar coincidencia exacta ignorando espacios y mayúsculas
          let coincidencia = opciones.find(opt => 
            opt.value && 
            opt.value.trim().toUpperCase() === valorEstado.trim().toUpperCase()
          );
          
          if (coincidencia) {
            estadoReservaSelect.value = coincidencia.value;
            console.log(`✅ Estado (coincidencia exacta): "${valorEstado}" -> "${estadoReservaSelect.value}"`);
          } else {
            // Intentar coincidencia parcial
            coincidencia = opciones.find(opt => 
              opt.value && valorEstado &&
              (opt.value.toUpperCase().includes(valorEstado.toUpperCase()) ||
              valorEstado.toUpperCase().includes(opt.value.toUpperCase()))
            );
            
            if (coincidencia) {
              estadoReservaSelect.value = coincidencia.value;
              console.log(`✅ Estado (coincidencia parcial): "${valorEstado}" -> "${estadoReservaSelect.value}"`);
            } else {
              // Intentar mapeo inteligente por palabras clave
              const valorUpper = valorEstado.toUpperCase();
              
              if (valorUpper.includes('CONFIRMAD') || valorUpper.includes('CONFIRM')) {
                coincidencia = opciones.find(opt => opt.value && opt.value.toUpperCase().includes('CONFIRMAD'));
              } else if (valorUpper.includes('PENDIENTE') || valorUpper.includes('PENDING')) {
                coincidencia = opciones.find(opt => opt.value && opt.value.toUpperCase().includes('PENDIENTE'));
              } else if (valorUpper.includes('CANCELAD') || valorUpper.includes('CANCEL')) {
                coincidencia = opciones.find(opt => opt.value && opt.value.toUpperCase().includes('CANCELAD'));
              }
              
              if (coincidencia) {
                estadoReservaSelect.value = coincidencia.value;
                console.log(`✅ Estado (mapeo inteligente): "${valorEstado}" -> "${estadoReservaSelect.value}"`);
              } else {
                console.error(`❌ No se encontró ninguna coincidencia para Estado: "${valorEstado}"`);
                console.log(`💡 Sugerencia: Verifica que el valor extraído coincida con alguna de estas opciones:`, opcionesDisponibles);
              }
            }
          }
        } else {
          console.log(`✅ Estado asignado correctamente: "${estadoReservaSelect.value}"`);
        }
      } else {
        console.log(`⚠️ Estado es null o inválido: "${valorEstado}"`);
      }
    }
    
    // Fecha de Viaje
    if (document.getElementById("fechaViaje")) {
      document.getElementById("fechaViaje").value = datosExtraidos.travelDate || "";
      console.log(`✅ Fecha Viaje: ${datosExtraidos.travelDate}`);
    }
    
    // Vendedor
    const vendedorSelect = document.getElementById("vendedor");
    if (vendedorSelect) {
      const valorVendedor = datosExtraidos.seller;
      if (valorVendedor && valorVendedor !== 'null' && valorVendedor !== null) {
        vendedorSelect.value = valorVendedor;
        console.log(`✅ Vendedor: "${valorVendedor}" -> "${vendedorSelect.value}"`);
        
        // Si no se seleccionó, intentar buscar una coincidencia parcial
        if (!vendedorSelect.value || vendedorSelect.value === "") {
          const opciones = Array.from(vendedorSelect.options);
          const coincidencia = opciones.find(opt => 
            opt.value && valorVendedor &&
            (opt.value.toUpperCase().includes(valorVendedor.toUpperCase()) ||
            valorVendedor.toUpperCase().includes(opt.value.toUpperCase()))
          );
          if (coincidencia) {
            vendedorSelect.value = coincidencia.value;
            console.log(`✅ Vendedor (coincidencia): "${valorVendedor}" -> "${vendedorSelect.value}"`);
          }
        }
      } else {
        console.log(`⚠️ Vendedor es null o inválido: "${valorVendedor}"`);
      }
    }
    
    // Cliente
    const clienteSelect = document.getElementById("cliente");
    if (clienteSelect) {
      const valorCliente = datosExtraidos.client;
      if (valorCliente && valorCliente !== 'null' && valorCliente !== null) {
        // Intentar asignación directa primero
        clienteSelect.value = valorCliente;
        
        // Si no se seleccionó, usar la función de coincidencia inteligente
        if (!clienteSelect.value || clienteSelect.value === "") {
          const clienteCoincidente = encontrarClienteCoincidente(valorCliente, masterData.clients);
          if (clienteCoincidente) {
            clienteSelect.value = clienteCoincidente;
            console.log(`✅ Cliente (coincidencia inteligente): "${valorCliente}" -> "${clienteSelect.value}"`);
          } else {
            console.warn(`⚠️ No se encontró coincidencia para Cliente: "${valorCliente}"`);
          }
        } else {
          console.log(`✅ Cliente asignado correctamente: "${clienteSelect.value}"`);
        }
      } else {
        console.log(`⚠️ Cliente es null o inválido: "${valorCliente}"`);
      }
    }
    
    // Campos adicionales de la reserva
    // Código
    if (document.getElementById("codigo")) {
      document.getElementById("codigo").value = datosExtraidos.codigo || "";
    }
    
    // Fecha de Reserva
    if (document.getElementById("fechaReserva")) {
      document.getElementById("fechaReserva").value = datosExtraidos.reservationDate || "";
    }
    
    // Fecha Fin de Tour
    if (document.getElementById("fechaFinTour")) {
      document.getElementById("fechaFinTour").value = datosExtraidos.tourEndDate || "";
    }
    
    // Fecha de Vencimiento
    if (document.getElementById("fechaVencimiento")) {
      document.getElementById("fechaVencimiento").value = datosExtraidos.dueDate || "";
    }
    
    // Contacto
    if (document.getElementById("contacto")) {
      document.getElementById("contacto").value = datosExtraidos.contact || "";
    }
    
    // Email de Contacto
    if (document.getElementById("contactEmail")) {
      document.getElementById("contactEmail").value = datosExtraidos.contactEmail || "";
    }
    
    // Teléfono de Contacto
    if (document.getElementById("contactPhone")) {
      document.getElementById("contactPhone").value = datosExtraidos.contactPhone || "";
    }
    
    // Moneda
    if (document.getElementById("moneda")) {
      document.getElementById("moneda").value = datosExtraidos.currency || "";
    }
    
    // Tipo de Cambio
    if (document.getElementById("tipoCambio")) {
      document.getElementById("tipoCambio").value = datosExtraidos.exchangeRate || "";
    }
    
    // Comisión
    if (document.getElementById("comision")) {
      document.getElementById("comision").value = datosExtraidos.commission || "";
    }
    
    // Monto Neto
    if (document.getElementById("montoNeto")) {
      document.getElementById("montoNeto").value = datosExtraidos.netAmount || "";
    }
    
    // Monto Bruto
    if (document.getElementById("montoBruto")) {
      document.getElementById("montoBruto").value = datosExtraidos.grossAmount || "";
    }
    
    // Nombre del Viaje
    if (document.getElementById("nombreViaje")) {
      document.getElementById("nombreViaje").value = datosExtraidos.tripName || "";
    }
    
    // Código de Producto
    if (document.getElementById("codigoProducto")) {
      document.getElementById("codigoProducto").value = datosExtraidos.productCode || "";
    }
    
    // Adultos
    if (document.getElementById("adultos")) {
      document.getElementById("adultos").value = datosExtraidos.adults || "";
    }
    
    // Niños
    if (document.getElementById("ninos")) {
      document.getElementById("ninos").value = datosExtraidos.children || "";
    }
    
    // Infantes
    if (document.getElementById("infantes")) {
      document.getElementById("infantes").value = datosExtraidos.infants || "";
    }
    
    // Proveedor
    if (document.getElementById("proveedor")) {
      document.getElementById("proveedor").value = datosExtraidos.provider || "";
    }
    
    // Código de Reserva
    if (document.getElementById("codigoReserva")) {
      document.getElementById("codigoReserva").value = datosExtraidos.reservationCode || "";
    }
    
    // Hotel (ahora es un objeto) - Solo mostrar si viene de la extracción
    const hotelSection = document.getElementById("hotelSection");
    const agregarHotelButton = document.getElementById("agregarHotel");
    const eliminarHotelButton = document.getElementById("eliminarHotel");
    
    if (datosExtraidos.hotel && typeof datosExtraidos.hotel === 'object') {
      // Mostrar sección de hotel y hacer campos required
      if (hotelSection) {
        hotelSection.style.display = "block";
      }
      if (agregarHotelButton) {
        agregarHotelButton.style.display = "none";
      }
      if (eliminarHotelButton) {
        eliminarHotelButton.style.display = "block";
      }
      
      // Llenar campos y hacerlos required
      const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
      camposHotel.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.required = true;
        }
      });
      
      if (document.getElementById("hotel_nombre")) {
        document.getElementById("hotel_nombre").value = datosExtraidos.hotel.nombre_hotel || "";
      }
      if (document.getElementById("hotel_tipo_habitacion")) {
        document.getElementById("hotel_tipo_habitacion").value = datosExtraidos.hotel.tipo_habitacion || "";
      }
      if (document.getElementById("hotel_ciudad")) {
        document.getElementById("hotel_ciudad").value = datosExtraidos.hotel.Ciudad || "";
      }
      if (document.getElementById("hotel_categoria")) {
        document.getElementById("hotel_categoria").value = datosExtraidos.hotel.Categoria || "";
      }
      if (document.getElementById("hotel_prioridad")) {
        let prioridadVal = datosExtraidos.hotel?.Prioridad || "";
        if (prioridadVal && !prioridadVal.startsWith("[")) {
          prioridadVal = "[" + prioridadVal + "]";
        }
        document.getElementById("hotel_prioridad").value = prioridadVal;
      }
      if (document.getElementById("hotel_in")) {
        document.getElementById("hotel_in").value = datosExtraidos.hotel.in || "";
      }
      if (document.getElementById("hotel_out")) {
        document.getElementById("hotel_out").value = datosExtraidos.hotel.out || "";
      }
    } else if (typeof datosExtraidos.hotel === 'string') {
      // Compatibilidad con formato antiguo (solo texto)
      if (hotelSection) {
        hotelSection.style.display = "block";
      }
      if (agregarHotelButton) {
        agregarHotelButton.style.display = "none";
      }
      if (eliminarHotelButton) {
        eliminarHotelButton.style.display = "block";
      }
      
      const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
      camposHotel.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.required = true;
        }
      });
      
      if (document.getElementById("hotel_nombre")) {
        document.getElementById("hotel_nombre").value = datosExtraidos.hotel || "";
      }
    } else {
      // No hay hotel - ocultar sección y mostrar botón
      if (hotelSection) {
        hotelSection.style.display = "none";
      }
      if (agregarHotelButton) {
        agregarHotelButton.style.display = "block";
      }
      if (eliminarHotelButton) {
        eliminarHotelButton.style.display = "none";
      }
      
      // Quitar required de campos de hotel
      const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
      camposHotel.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.required = false;
          campo.value = "";
        }
      });
    }
    
    // Check In
    if (document.getElementById("checkIn")) {
      document.getElementById("checkIn").value = datosExtraidos.checkIn || "";
    }
    
    // Check Out
    if (document.getElementById("checkOut")) {
      document.getElementById("checkOut").value = datosExtraidos.checkOut || "";
    }
    
    // Estado de Deuda
    if (document.getElementById("estadoDeuda")) {
      document.getElementById("estadoDeuda").value = datosExtraidos.estadoDeuda || "";
    }
    
    // Llenar servicios si existen
    const serviciosSection = document.getElementById("serviciosSection");
    const agregarServicioButton = document.getElementById("agregarServicio");
    
    if (datosExtraidos.services && Array.isArray(datosExtraidos.services) && datosExtraidos.services.length > 0) {
      if (serviciosSection) {
        serviciosSection.style.display = "block";
      }
      llenarServicios(datosExtraidos.services);
    } else {
      // Si no hay servicios, ocultar la sección pero mantener el botón visible
      if (serviciosSection) {
        serviciosSection.style.display = "none";
      }
    }
    
    // El botón de agregar servicio siempre está visible (definido en HTML)
    
    // Llenar vuelos si existen
    if (datosExtraidos.flights && Array.isArray(datosExtraidos.flights) && datosExtraidos.flights.length > 0) {
      const vuelosSection = document.getElementById("vuelosSection");
      if (vuelosSection) {
        vuelosSection.style.display = "block";
      }
      llenarVuelos(datosExtraidos.flights);
    }
    
    // Actualizar estado del botón después de llenar los datos
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  }, 250); // Esperar 250ms para asegurar que los selects estén poblados
}

/**
 * Encuentra el mejor servicio coincidente usando regex y comparación de palabras
 * @param {string} servicioBackend - Nombre del servicio que viene del backend
 * @returns {string|null} - El servicio más coincidente o null si no hay coincidencias suficientes
 */
function encontrarServicioCoincidente(servicioBackend) {
  if (!servicioBackend || typeof servicioBackend !== 'string') {
    return null;
  }
  
  const servicioNormalizado = servicioBackend.trim().toUpperCase();
  const palabrasBackend = servicioNormalizado.split(/\s+/).filter(p => p.length > 0);
  
  if (palabrasBackend.length === 0) {
    return null;
  }
  
  let mejorCoincidencia = null;
  let maxCoincidencias = 0;
  let mejorPuntuacion = 0;
  
  // Palabras clave importantes que dan más peso a la coincidencia
  const palabrasClave = ['BODEGA', 'BODEGAS', 'TRASLADO', 'EXCURSION', 'TOUR', 'ALMUERZO', 'CENA', 'HOTEL', 'SPA', 'CABALGATA'];
  
  servicesList.forEach(servicioLista => {
    const servicioListaNormalizado = servicioLista.trim().toUpperCase();
    const palabrasLista = servicioListaNormalizado.split(/\s+/).filter(p => p.length > 0);
    
    // Contar palabras que coinciden
    let coincidencias = 0;
    let puntuacion = 0;
    
    palabrasBackend.forEach(palabraBackend => {
      // Buscar la palabra en la lista (coincidencia exacta o parcial)
      const encontrado = palabrasLista.some(palabraLista => {
        // Coincidencia exacta
        if (palabraLista === palabraBackend) {
          // Dar más peso si es una palabra clave
          if (palabrasClave.includes(palabraBackend)) {
            puntuacion += 2;
          } else {
            puntuacion += 1;
          }
          return true;
        }
        // Coincidencia parcial (una contiene a la otra)
        if (palabraLista.includes(palabraBackend) || palabraBackend.includes(palabraLista)) {
          // Solo contar si ambas palabras tienen al menos 3 caracteres para evitar falsos positivos
          if (palabraBackend.length >= 3 && palabraLista.length >= 3) {
            // Dar peso reducido a coincidencias parciales
            if (palabrasClave.includes(palabraBackend) || palabrasClave.includes(palabraLista)) {
              puntuacion += 1;
            } else {
              puntuacion += 0.5;
            }
            return true;
          }
        }
        return false;
      });
      
      if (encontrado) {
        coincidencias++;
      }
    });
    
    // Si hay más de una palabra coincidente, considerar esta opción
    // O si hay una coincidencia muy significativa (palabra clave con alta puntuación)
    if (coincidencias > maxCoincidencias && (coincidencias > 1 || puntuacion >= 2)) {
      maxCoincidencias = coincidencias;
      mejorPuntuacion = puntuacion;
      mejorCoincidencia = servicioLista;
    } else if (coincidencias === maxCoincidencias && puntuacion > mejorPuntuacion) {
      // Si hay el mismo número de coincidencias, elegir el de mayor puntuación
      mejorPuntuacion = puntuacion;
      mejorCoincidencia = servicioLista;
    }
  });
  
  // Si encontramos al menos 2 palabras coincidentes, o una coincidencia muy significativa
  if (maxCoincidencias >= 2 || mejorPuntuacion >= 2) {
    return mejorCoincidencia;
  }
  
  return null;
}

/**
 * Normaliza el estado del servicio a su código (ej: "LI - LIBERADO [LI]" -> "LI")
 * @param {string} estado - Estado del servicio (puede venir en diferentes formatos)
 * @returns {string} Código del estado o string vacío
 */
function normalizarEstadoServicio(estado) {
  if (!estado || typeof estado !== 'string') return '';
  
  // Si ya es un código de 2 letras, retornarlo
  if (estado.length === 2 && /^[A-Z]{2}$/.test(estado)) {
    return estado;
  }
  
  // Intentar extraer el código del formato "XX - DESCRIPCION [XX]"
  const match = estado.match(/^([A-Z]{2})\s*-/);
  if (match) {
    return match[1];
  }
  
  // Intentar extraer de formato "[XX]"
  const match2 = estado.match(/\[([A-Z]{2})\]/);
  if (match2) {
    return match2[1];
  }
  
  return '';
}

/**
 * Crea un elemento de servicio en el formulario
 * @param {number} index - Índice del servicio
 * @param {Object} servicio - Datos del servicio (opcional)
 * @param {boolean} esDeExtraccion - Si el servicio viene de la extracción de IA
 * @returns {HTMLElement} El elemento div del servicio
 */
function crearElementoServicio(index, servicio = {}, esDeExtraccion = false) {
  const servicioDiv = document.createElement("div");
  servicioDiv.className = "servicio-item";
  
  // Intentar encontrar el servicio coincidente
  const servicioCoincidente = servicio.servicio 
    ? encontrarServicioCoincidente(servicio.servicio)
    : null;
  
  // Normalizar el estado del servicio
  const estadoNormalizado = normalizarEstadoServicio(servicio.estado || '');
  const categoriaNormalizada = (servicio.categoria || servicio.Categoria || '').trim();
  let prioridadNormalizada = (servicio.prioridad || servicio.Prioridad || '').trim();
  if (prioridadNormalizada && !prioridadNormalizada.startsWith('[')) {
    prioridadNormalizada = '[' + prioridadNormalizada + ']';
  }
  
  // Crear el HTML con un datalist para autocompletado
  const datalistId = `serviciosList_${index}`;
  const requiredAttr = esDeExtraccion ? 'required' : '';
  const requiredSpan = esDeExtraccion ? '<span style="color: red;">*</span>' : '';
  
  servicioDiv.dataset.servicioIndex = index;
  servicioDiv.innerHTML = `
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
      <h4 style="margin: 0;">Servicio ${index + 1}</h4>
      <button class="btn-eliminar-servicio" title="Eliminar servicio" data-servicio-index="${index}">✕</button>
    </div>
    <div class="form-group">
      <label>Destino: ${requiredSpan}</label>
      <input type="text" id="servicio_destino_${index}" value="${servicio.destino || ''}" ${requiredAttr}>
    </div>
    <div class="form-group">
      <label>Fecha Entrada: ${requiredSpan}</label>
      <input type="date" id="servicio_in_${index}" value="${servicio.in || ''}" ${requiredAttr}>
    </div>
    <div class="form-group">
      <label>Fecha Salida: ${requiredSpan}</label>
      <input type="date" id="servicio_out_${index}" value="${servicio.out || ''}" ${requiredAttr}>
    </div>
    <div class="form-group">
      <label>Noches: ${requiredSpan}</label>
      <input type="number" id="servicio_nts_${index}" value="${servicio.nts !== undefined && servicio.nts !== null ? servicio.nts : ''}" min="0" ${requiredAttr}>
    </div>
    <div class="form-group">
      <label>Base Pax: ${requiredSpan}</label>
      <input type="number" id="servicio_basePax_${index}" value="${servicio.basePax || ''}" ${requiredAttr}>
    </div>
    <div class="form-group">
      <label>Servicio: ${requiredSpan}</label>
      <input type="text" id="servicio_servicio_${index}" 
             list="${datalistId}" 
             value="${servicioCoincidente || servicio.servicio || ''}" 
             placeholder="Buscar o escribir servicio..."
             ${requiredAttr}>
      <datalist id="${datalistId}">
        ${servicesList.map(s => `<option class="servicio-option" value="${s}">${s}</option>`).join('')}
      </datalist>
    </div>
    <div class="form-group">
      <label>Descripción:</label>
      <textarea id="servicio_descripcion_${index}">${servicio.descripcion || ''}</textarea>
    </div>
    <div class="form-group">
      <label>Estado: ${requiredSpan}</label>
      <select id="servicio_estado_${index}" ${requiredAttr}>
        <option value="">Seleccione...</option>
        <option value="LI" ${estadoNormalizado === 'LI' ? 'selected' : ''}>LI - LIBERADO [LI]</option>
        <option value="OK" ${estadoNormalizado === 'OK' ? 'selected' : ''}>OK - CONFIRMADO [OK]</option>
        <option value="WL" ${estadoNormalizado === 'WL' ? 'selected' : ''}>WL - LISTA DE ESPERA [WL]</option>
        <option value="RM" ${estadoNormalizado === 'RM' ? 'selected' : ''}>RM - FAVOR MODIFICAR [RM]</option>
        <option value="NN" ${estadoNormalizado === 'NN' ? 'selected' : ''}>NN - FAVOR RESERVAR [NN]</option>
        <option value="RQ" ${estadoNormalizado === 'RQ' ? 'selected' : ''}>RQ - REQUERIDO [RQ]</option>
        <option value="LK" ${estadoNormalizado === 'LK' ? 'selected' : ''}>LK - RVA OK S/LIQUIDAR [LK]</option>
        <option value="RE" ${estadoNormalizado === 'RE' ? 'selected' : ''}>RE - RECHAZADO [RE]</option>
        <option value="MQ" ${estadoNormalizado === 'MQ' ? 'selected' : ''}>MQ - MODIFICACION REQUERIDA [MQ]</option>
        <option value="CL" ${estadoNormalizado === 'CL' ? 'selected' : ''}>CL - FAVOR CANCELAR [CL]</option>
        <option value="CA" ${estadoNormalizado === 'CA' ? 'selected' : ''}>CA - CANCELACION SOLICITADA [CA]</option>
        <option value="CX" ${estadoNormalizado === 'CX' ? 'selected' : ''}>CX - CANCELADO [CX]</option>
        <option value="EM" ${estadoNormalizado === 'EM' ? 'selected' : ''}>EM - EMITIDO [EM]</option>
        <option value="EN" ${estadoNormalizado === 'EN' ? 'selected' : ''}>EN - ENTREGADO [EN]</option>
        <option value="AR" ${estadoNormalizado === 'AR' ? 'selected' : ''}>AR - FAVOR RESERVAR [AR]</option>
        <option value="HK" ${estadoNormalizado === 'HK' ? 'selected' : ''}>HK - OK CUPO [HK]</option>
        <option value="PE" ${estadoNormalizado === 'PE' ? 'selected' : ''}>PE - PENALIDAD [PE]</option>
        <option value="NO" ${estadoNormalizado === 'NO' ? 'selected' : ''}>NO - NEGADO [NO]</option>
        <option value="NC" ${estadoNormalizado === 'NC' ? 'selected' : ''}>NC - NO CONFORMIDAD [NC]</option>
        <option value="PF" ${estadoNormalizado === 'PF' ? 'selected' : ''}>PF - PENDIENTE DE FC. COMISION [PF]</option>
        <option value="AO" ${estadoNormalizado === 'AO' ? 'selected' : ''}>AO - REQUERIR ON LINE [AO]</option>
        <option value="CO" ${estadoNormalizado === 'CO' ? 'selected' : ''}>CO - CANCELAR ONLINE [CO]</option>
        <option value="GX" ${estadoNormalizado === 'GX' ? 'selected' : ''}>GX - GASTOS CANCELACION ONLINE [GX]</option>
        <option value="EO" ${estadoNormalizado === 'EO' ? 'selected' : ''}>EO - EN TRAFICO [EO]</option>
        <option value="KL" ${estadoNormalizado === 'KL' ? 'selected' : ''}>KL - REQUERIDO CUPO [KL]</option>
        <option value="MI" ${estadoNormalizado === 'MI' ? 'selected' : ''}>MI - RESERVA MIGRADA [MI]</option>
        <option value="VO" ${estadoNormalizado === 'VO' ? 'selected' : ''}>VO - VOID [VO]</option>
      </select>
    </div>
    <div class="form-group">
      <label>Categoría: <span style="color: #ef4444;">*</span></label>
      <select id="servicio_categoria_${index}" required>
        <option value="">Seleccione...</option>
        <option value="Regular" ${categoriaNormalizada === 'Regular' ? 'selected' : ''}>Regular</option>
        <option value="Privado" ${categoriaNormalizada === 'Privado' ? 'selected' : ''}>Privado</option>
      </select>
    </div>
    <div class="form-group">
      <label>Prioridad: <span style="color: #ef4444;">*</span></label>
      <select id="servicio_prioridad_${index}" required>
        <option value="">Seleccione...</option>
        <option value="[1]" ${prioridadNormalizada === '[1]' ? 'selected' : ''}>Unica [1]</option>
        <option value="[NA]" ${prioridadNormalizada === '[NA]' ? 'selected' : ''}>NACIONAL [NA]</option>
        <option value="[EX]" ${prioridadNormalizada === '[EX]' ? 'selected' : ''}>EXTRANJERO [EX]</option>
        <option value="[DE]" ${prioridadNormalizada === '[DE]' ? 'selected' : ''}>DESPEGAR [DE]</option>
        <option value="[VD]" ${prioridadNormalizada === '[VD]' ? 'selected' : ''}>VENTA DIRECTA [VD]</option>
        <option value="[NE]" ${prioridadNormalizada === '[NE]' ? 'selected' : ''}>NACIONAL EXTRANJERO [NE]</option>
        <option value="[AZUL]" ${prioridadNormalizada === '[AZUL]' ? 'selected' : ''}>AZUL VIAGENS [AZUL]</option>
        <option value="[AZ]" ${prioridadNormalizada === '[AZ]' ? 'selected' : ''}>AZUL VIAGENS [AZ]</option>
        <option value="[CV]" ${prioridadNormalizada === '[CV]' ? 'selected' : ''}>CVC [CV]</option>
      </select>
    </div>
  `;
  
  // Agregar event listeners para validación en tiempo real
  setTimeout(() => {
    const camposServicio = [
      `servicio_destino_${index}`,
      `servicio_in_${index}`,
      `servicio_out_${index}`,
      `servicio_nts_${index}`,
      `servicio_basePax_${index}`,
      `servicio_servicio_${index}`,
      `servicio_descripcion_${index}`,
      `servicio_estado_${index}`,
      `servicio_categoria_${index}`,
      `servicio_prioridad_${index}`
    ];
    
    camposServicio.forEach(campoId => {
      const campo = document.getElementById(campoId);
      if (campo) {
        campo.addEventListener('change', actualizarEstadoBotonCrearReserva);
        campo.addEventListener('input', actualizarEstadoBotonCrearReserva);
      }
    });
    
    // Agregar event listener al botón de eliminar
    const btnEliminar = servicioDiv.querySelector('.btn-eliminar-servicio');
    if (btnEliminar) {
      btnEliminar.onclick = function(e) {
        e.stopPropagation();
        eliminarServicio(servicioDiv);
      };
    }
  }, 100);
  
  return servicioDiv;
}

/**
 * Llena los servicios en el formulario
 * @param {Array} servicios - Array de servicios extraídos
 */
function llenarServicios(servicios) {
  const serviciosContainer = document.getElementById("serviciosContainer");
  const serviciosSection = document.getElementById("serviciosSection");
  if (!serviciosContainer) return;
  
  serviciosContainer.innerHTML = "";
  
  // Mostrar sección de servicios
  if (serviciosSection) {
    serviciosSection.style.display = "block";
  }
  
  servicios.forEach((servicio, index) => {
    const servicioDiv = crearElementoServicio(index, servicio, true);
    serviciosContainer.appendChild(servicioDiv);
  });
  
  // Actualizar estado del botón después de crear los servicios
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

/**
 * Agrega un nuevo servicio manualmente
 */
function agregarNuevoServicio() {
  const serviciosContainer = document.getElementById("serviciosContainer");
  const serviciosSection = document.getElementById("serviciosSection");
  
  if (!serviciosContainer) return;
  
  // Mostrar sección si está oculta
  if (serviciosSection) {
    serviciosSection.style.display = "block";
  }
  
  // Contar servicios existentes
  const servicioItems = serviciosContainer.querySelectorAll(".servicio-item");
  const nuevoIndex = servicioItems.length;
  
  // Crear nuevo servicio (sin required porque es manual)
  const nuevoServicio = crearElementoServicio(nuevoIndex, {}, false);
  serviciosContainer.appendChild(nuevoServicio);
  
  // Scroll suave hacia el nuevo servicio
  if (nuevoServicio.scrollIntoView) {
    nuevoServicio.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  
  // Actualizar estado del botón
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

/**
 * Muestra la sección de hotel y hace los campos required
 */
function mostrarSeccionHotel() {
  const hotelSection = document.getElementById("hotelSection");
  const agregarHotelButton = document.getElementById("agregarHotel");
  const eliminarHotelButton = document.getElementById("eliminarHotel");
  
  if (hotelSection) {
    hotelSection.style.display = "block";
  }
  
  if (agregarHotelButton) {
    agregarHotelButton.style.display = "none";
  }
  
  // Mostrar botón de eliminar hotel
  if (eliminarHotelButton) {
    eliminarHotelButton.style.display = "block";
  }
  
  // Hacer campos required cuando se agrega manualmente
  const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
  camposHotel.forEach(campoId => {
    const campo = document.getElementById(campoId);
    if (campo) {
      campo.required = true;
    }
  });
  
  // Actualizar estado del botón
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

/**
 * Elimina un servicio del formulario
 * @param {HTMLElement} servicioDiv - El elemento div del servicio a eliminar
 */
function eliminarServicio(servicioDiv) {
  try {
    const serviciosContainer = document.getElementById("serviciosContainer");
    if (!serviciosContainer) return;
    
    // Eliminar el servicio directamente sin restricciones
    servicioDiv.remove();
    
    // Renumerar servicios
    renumerarServicios();
    
    // Verificar si quedan servicios después de eliminar
    const servicioItems = serviciosContainer.querySelectorAll(".servicio-item");
    if (servicioItems.length === 0) {
      // Si no hay servicios, ocultar la sección y limpiar el contenedor
      const serviciosSection = document.getElementById("serviciosSection");
      if (serviciosSection) {
        serviciosSection.style.display = "none";
      }
      serviciosContainer.innerHTML = "";
    }
    
    mostrarMensaje("Servicio eliminado correctamente", "success");
    
    // Actualizar estado del botón después de eliminar
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  } catch (error) {
    mostrarMensaje("Error al eliminar servicio", "error");
  }
}

/**
 * Renumera los servicios después de eliminar uno
 */
function renumerarServicios() {
  const serviciosContainer = document.getElementById("serviciosContainer");
  if (!serviciosContainer) return;
  
  const servicioItems = Array.from(serviciosContainer.querySelectorAll(".servicio-item"));
  
  servicioItems.forEach((servicioDiv, index) => {
    const nuevoIndex = index;
    
    // Actualizar el dataset
    servicioDiv.dataset.servicioIndex = nuevoIndex;
    
    // Actualizar el título
    const tituloDiv = servicioDiv.querySelector("div:first-child");
    if (tituloDiv) {
      const titulo = tituloDiv.querySelector("h4");
      if (titulo) {
        titulo.textContent = `Servicio ${nuevoIndex + 1}`;
      }
    }
    
    // Actualizar IDs de todos los campos - buscar por patrón dentro del div
    const campos = [
      'servicio_destino',
      'servicio_in',
      'servicio_out',
      'servicio_nts',
      'servicio_basePax',
      'servicio_servicio',
      'servicio_descripcion',
      'servicio_estado',
      'servicio_categoria',
      'servicio_prioridad'
    ];
    
    campos.forEach(campoBase => {
      // Buscar el campo dentro del servicioDiv por el patrón del ID
      const campoViejo = servicioDiv.querySelector(`[id^="${campoBase}_"]`);
      if (campoViejo) {
        // Guardar el valor actual antes de cambiar el ID (especialmente para select)
        const valorActual = campoViejo.value;
        
        campoViejo.id = `${campoBase}_${nuevoIndex}`;
        
        // Restaurar el valor después de cambiar el ID
        campoViejo.value = valorActual;
        
        // Actualizar el list del datalist si es el campo servicio
        if (campoBase === 'servicio_servicio') {
          const datalistId = `serviciosList_${nuevoIndex}`;
          campoViejo.setAttribute('list', datalistId);
          
          // Actualizar el datalist
          const datalist = servicioDiv.querySelector('datalist');
          if (datalist) {
            datalist.id = datalistId;
          }
        }
      }
    });
    
    // Actualizar el data-servicio-index del botón eliminar
    const btnEliminar = servicioDiv.querySelector('.btn-eliminar-servicio');
    if (btnEliminar) {
      btnEliminar.setAttribute('data-servicio-index', nuevoIndex);
    }
  });
}

/**
 * Crea un elemento de vuelo en el formulario
 * @param {number} index - Índice del vuelo
 * @param {Object} vuelo - Datos del vuelo (opcional)
 * @param {boolean} esDeExtraccion - Si el vuelo viene de la extracción de IA
 * @returns {HTMLElement} El elemento div del vuelo
 */
function crearElementoVuelo(index, vuelo = {}, esDeExtraccion = false) {
  // Mapear campos del backend a los campos del formulario
  // El backend envía: departureDate, departureTime, arrivalDate, arrivalTime
  // También mantener compatibilidad con formato antiguo: date, time
  const departureDate = vuelo.departureDate || vuelo.date || '';
  const departureTime = vuelo.departureTime || vuelo.time || '';
  const arrivalDate = vuelo.arrivalDate || '';
  const arrivalTime = vuelo.arrivalTime || '';
  
  const vueloDiv = document.createElement("div");
  vueloDiv.className = "vuelo-item";
  vueloDiv.dataset.vueloIndex = index;
  
  // Los campos de vuelos siempre son editables
  vueloDiv.innerHTML = `
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
      <h4 style="margin: 0;">Vuelo ${index + 1}</h4>
      <button class="btn-eliminar-vuelo" title="Eliminar vuelo" data-vuelo-index="${index}">✕</button>
    </div>
    <div class="form-group">
      <label>Aerolínea:</label>
      <input type="text" id="vuelo_airline_${index}" value="${vuelo.airline || ''}">
    </div>
    <div class="form-group">
      <label>Número de Vuelo:</label>
      <input type="text" id="vuelo_flightNumber_${index}" value="${vuelo.flightNumber || ''}">
    </div>
    <div class="form-group">
      <label>Origen:</label>
      <input type="text" id="vuelo_origin_${index}" value="${vuelo.origin || ''}">
    </div>
    <div class="form-group">
      <label>Destino:</label>
      <input type="text" id="vuelo_destination_${index}" value="${vuelo.destination || ''}">
    </div>
    <div class="form-group">
      <label>Fecha de Salida:</label>
      <input type="date" id="vuelo_departureDate_${index}" value="${departureDate}">
    </div>
    <div class="form-group">
      <label>Hora de Salida:</label>
      <input type="time" id="vuelo_departureTime_${index}" value="${departureTime}">
    </div>
    <div class="form-group">
      <label>Fecha de Llegada:</label>
      <input type="date" id="vuelo_arrivalDate_${index}" value="${arrivalDate}">
    </div>
    <div class="form-group">
      <label>Hora de Llegada:</label>
      <input type="time" id="vuelo_arrivalTime_${index}" value="${arrivalTime}">
    </div>
  `;
  
  // Agregar event listener al botón de eliminar
  setTimeout(() => {
    const btnEliminar = vueloDiv.querySelector('.btn-eliminar-vuelo');
    if (btnEliminar) {
      btnEliminar.onclick = function(e) {
        e.stopPropagation();
        eliminarVuelo(vueloDiv);
      };
    }
  }, 100);
  
  return vueloDiv;
}

/**
 * Llena los vuelos en el formulario
 * @param {Array} vuelos - Array de vuelos extraídos
 */
function llenarVuelos(vuelos) {
  const vuelosContainer = document.getElementById("vuelosContainer");
  const vuelosSection = document.getElementById("vuelosSection");
  if (!vuelosContainer) return;
  
  vuelosContainer.innerHTML = "";
  
  // Mostrar sección de vuelos
  if (vuelosSection) {
    vuelosSection.style.display = "block";
  }
  
  // Mostrar botón de agregar vuelo
  const agregarVueloButton = document.getElementById("agregarVuelo");
  if (agregarVueloButton) {
    agregarVueloButton.style.display = "block";
  }
  
  vuelos.forEach((vuelo, index) => {
    const vueloDiv = crearElementoVuelo(index, vuelo, true);
    vuelosContainer.appendChild(vueloDiv);
  });
  
  // Actualizar estado del botón después de crear los vuelos
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

/**
 * Agrega un nuevo vuelo manualmente
 */
function agregarNuevoVuelo() {
  const vuelosContainer = document.getElementById("vuelosContainer");
  const vuelosSection = document.getElementById("vuelosSection");
  
  if (!vuelosContainer) return;
  
  // Mostrar sección si está oculta
  if (vuelosSection) {
    vuelosSection.style.display = "block";
  }
  
  // Mostrar botón de agregar vuelo
  const agregarVueloButton = document.getElementById("agregarVuelo");
  if (agregarVueloButton) {
    agregarVueloButton.style.display = "block";
  }
  
  // Contar vuelos existentes
  const vueloItems = vuelosContainer.querySelectorAll(".vuelo-item");
  const nuevoIndex = vueloItems.length;
  
  // Crear nuevo vuelo (sin readonly porque es manual)
  const nuevoVuelo = crearElementoVuelo(nuevoIndex, {}, false);
  vuelosContainer.appendChild(nuevoVuelo);
  
  // Scroll suave hacia el nuevo vuelo
  if (nuevoVuelo.scrollIntoView) {
    nuevoVuelo.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  
  // Actualizar estado del botón
  setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
}

/**
 * Elimina un vuelo del formulario
 * @param {HTMLElement} vueloDiv - El elemento div del vuelo a eliminar
 */
function eliminarVuelo(vueloDiv) {
  try {
    const vuelosContainer = document.getElementById("vuelosContainer");
    if (!vuelosContainer) return;
    
    // Eliminar el vuelo directamente sin restricciones
    vueloDiv.remove();
    
    // Renumerar vuelos
    renumerarVuelos();
    
    // Verificar si quedan vuelos después de eliminar
    const vueloItems = vuelosContainer.querySelectorAll(".vuelo-item");
    if (vueloItems.length === 0) {
      // Si no hay vuelos, ocultar la sección y limpiar el contenedor
      const vuelosSection = document.getElementById("vuelosSection");
      if (vuelosSection) {
        vuelosSection.style.display = "none";
      }
      vuelosContainer.innerHTML = "";
    }
    
    mostrarMensaje("Vuelo eliminado correctamente", "success");
    
    // Actualizar estado del botón después de eliminar
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  } catch (error) {
    mostrarMensaje("Error al eliminar vuelo", "error");
  }
}

/**
 * Renumera los vuelos después de eliminar uno
 */
function renumerarVuelos() {
  const vuelosContainer = document.getElementById("vuelosContainer");
  if (!vuelosContainer) return;
  
  const vueloItems = Array.from(vuelosContainer.querySelectorAll(".vuelo-item"));
  
  vueloItems.forEach((vueloDiv, index) => {
    const nuevoIndex = index;
    
    // Actualizar el dataset
    vueloDiv.dataset.vueloIndex = nuevoIndex;
    
    // Actualizar el título
    const tituloDiv = vueloDiv.querySelector("div:first-child");
    if (tituloDiv) {
      const titulo = tituloDiv.querySelector("h4");
      if (titulo) {
        titulo.textContent = `Vuelo ${nuevoIndex + 1}`;
      }
    }
    
    // Actualizar IDs de todos los campos
    const campos = [
      'vuelo_airline',
      'vuelo_flightNumber',
      'vuelo_origin',
      'vuelo_destination',
      'vuelo_departureDate',
      'vuelo_departureTime',
      'vuelo_arrivalDate',
      'vuelo_arrivalTime'
    ];
    
    campos.forEach(campoBase => {
      const campoViejo = vueloDiv.querySelector(`[id^="${campoBase}_"]`);
      if (campoViejo) {
        const valorActual = campoViejo.value;
        campoViejo.id = `${campoBase}_${nuevoIndex}`;
        campoViejo.value = valorActual;
      }
    });
    
    // Actualizar el data-vuelo-index del botón eliminar
    const btnEliminar = vueloDiv.querySelector('.btn-eliminar-vuelo');
    if (btnEliminar) {
      btnEliminar.setAttribute('data-vuelo-index', nuevoIndex);
    }
  });
}

/**
 * Elimina el hotel del formulario
 */
function eliminarHotel() {
  try {
    const hotelSection = document.getElementById("hotelSection");
    const agregarHotelButton = document.getElementById("agregarHotel");
    
    if (!hotelSection) return;
    
    // Limpiar campos del hotel
    const camposHotel = ['hotel_nombre', 'hotel_tipo_habitacion', 'hotel_ciudad', 'hotel_categoria', 'hotel_prioridad', 'hotel_in', 'hotel_out'];
    camposHotel.forEach(campoId => {
      const campo = document.getElementById(campoId);
      if (campo) {
        campo.value = "";
        campo.required = false;
      }
    });
    
    // Ocultar sección de hotel
    hotelSection.style.display = "none";
    
    // Mostrar botón de agregar hotel
    if (agregarHotelButton) {
      agregarHotelButton.style.display = "block";
    }
    
    // Ocultar botón de eliminar hotel
    const eliminarHotelButton = document.getElementById("eliminarHotel");
    if (eliminarHotelButton) {
      eliminarHotelButton.style.display = "none";
    }
    
    mostrarMensaje("Hotel eliminado correctamente", "success");
    
    // Actualizar estado del botón después de eliminar
    setTimeout(() => actualizarEstadoBotonCrearReserva(), 200);
  } catch (error) {
    mostrarMensaje("Error al eliminar hotel", "error");
  }
}

/**
 * Valida si todos los campos obligatorios están completos
 * @returns {boolean} true si todos los campos están completos
 */
function validarCamposObligatorios() {
  const container = document.getElementById("pasajerosContainer");
  if (!container) return false;
  
  const pasajeros = container.querySelectorAll(".pasajero-acordeon");
  if (pasajeros.length === 0) return false;
  
  // Validar pasajeros (solo nombre y apellido)
  let pasajerosValidos = true;
  pasajeros.forEach((pasajeroDiv) => {
    const numero = pasajeroDiv.dataset.numeroPasajero;
    const content = pasajeroDiv.querySelector(".pasajero-content");
    
    if (content) {
      const nombre = content.querySelector(`#nombre_${numero}`)?.value || "";
      const apellido = content.querySelector(`#apellido_${numero}`)?.value || "";
      
      if (nombre.trim() === "" || apellido.trim() === "") {
        pasajerosValidos = false;
      }
    }
  });
  
  // Validar datos de reserva (todos obligatorios)
  const tipoReserva = document.getElementById("tipoReserva")?.value || "";
  const estadoReserva = document.getElementById("estadoReserva")?.value || "";
  const fechaViaje = document.getElementById("fechaViaje")?.value || "";
  const vendedor = document.getElementById("vendedor")?.value || "";
  const cliente = document.getElementById("cliente")?.value || "";
  const moneda = document.getElementById("moneda")?.value || "";
  
  const reservaValida = 
    tipoReserva.trim() !== "" &&
    estadoReserva.trim() !== "" &&
    fechaViaje.trim() !== "" &&
    vendedor.trim() !== "" &&
    cliente.trim() !== "" &&
    moneda.trim() !== "";
  
  // Validar campos obligatorios del hotel (solo si la sección está visible)
  const hotelSection = document.getElementById("hotelSection");
  let hotelValido = true;
  
  if (hotelSection && hotelSection.style.display !== "none") {
    const hotelNombre = document.getElementById("hotel_nombre");
    const hotelTipoHabitacion = document.getElementById("hotel_tipo_habitacion");
    const hotelCiudad = document.getElementById("hotel_ciudad");
    const hotelCategoria = document.getElementById("hotel_categoria");
    const hotelPrioridad = document.getElementById("hotel_prioridad");
    const hotelIn = document.getElementById("hotel_in");
    const hotelOut = document.getElementById("hotel_out");
    
    // Solo validar si los campos son required
    if (hotelNombre?.required && (!hotelNombre.value || hotelNombre.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelTipoHabitacion?.required && (!hotelTipoHabitacion.value || hotelTipoHabitacion.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelCiudad?.required && (!hotelCiudad.value || hotelCiudad.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelCategoria?.required && (!hotelCategoria.value || hotelCategoria.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelPrioridad?.required && (!hotelPrioridad.value || hotelPrioridad.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelIn?.required && (!hotelIn.value || hotelIn.value.trim() === "")) {
      hotelValido = false;
    }
    if (hotelOut?.required && (!hotelOut.value || hotelOut.value.trim() === "")) {
      hotelValido = false;
    }
  }
  
  // Validar campos obligatorios de servicios (solo si la sección está visible)
  const serviciosSection = document.getElementById("serviciosSection");
  const serviciosContainer = document.getElementById("serviciosContainer");
  let serviciosValidos = true;
  
  if (serviciosSection && serviciosSection.style.display !== "none" && serviciosContainer) {
    const servicioItems = serviciosContainer.querySelectorAll(".servicio-item");
    
    if (servicioItems.length > 0) {
      servicioItems.forEach((item, index) => {
        const destino = document.getElementById(`servicio_destino_${index}`);
        const servicioIn = document.getElementById(`servicio_in_${index}`);
        const servicioOut = document.getElementById(`servicio_out_${index}`);
        const nts = document.getElementById(`servicio_nts_${index}`);
        const basePax = document.getElementById(`servicio_basePax_${index}`);
        const servicio = document.getElementById(`servicio_servicio_${index}`);
        const estado = document.getElementById(`servicio_estado_${index}`);
        const servicioCategoria = document.getElementById(`servicio_categoria_${index}`);
        const servicioPrioridad = document.getElementById(`servicio_prioridad_${index}`);
        
        // Solo validar si los campos son required
        if (destino?.required && (!destino.value || destino.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (servicioIn?.required && (!servicioIn.value || servicioIn.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (servicioOut?.required && (!servicioOut.value || servicioOut.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (nts?.required && (nts.value === null || nts.value === undefined || nts.value === "")) {
          serviciosValidos = false;
        }
        if (basePax?.required && (!basePax.value || basePax.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (servicio?.required && (!servicio.value || servicio.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (estado?.required && (!estado.value || estado.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (servicioCategoria?.required && (!servicioCategoria.value || servicioCategoria.value.trim() === "")) {
          serviciosValidos = false;
        }
        if (servicioPrioridad?.required && (!servicioPrioridad.value || servicioPrioridad.value.trim() === "")) {
          serviciosValidos = false;
        }
      });
    }
  }
  
  return pasajerosValidos && reservaValida && hotelValido && serviciosValidos;
}

/**
 * Actualiza el texto del botón según si es crear o editar
 * @param {boolean} isEdit - Si es true, muestra "Editar Reserva"; si es false, muestra "Crear Reserva"
 */
function actualizarTextoBotonReserva(isEdit = false) {
  const boton = document.getElementById("crearReserva");
  if (!boton) return;
  
  const label = boton.querySelector('.ms-Button-label');
  if (label) {
    const texto = isEdit 
      ? "✏️ Editar Reserva en iTraffic" 
      : "🚀 Crear Reserva en iTraffic";
    label.textContent = texto;
    console.log('📝 Texto del botón actualizado:', texto, 'isEdit:', isEdit);
  }
}

/**
 * Actualiza el estado del botón "Crear Reserva" según la validación
 */
function actualizarEstadoBotonCrearReserva() {
  const boton = document.getElementById("crearReserva");
  if (!boton) return;
  
  const esValido = validarCamposObligatorios();
  const doesReservationExist = extractionState.doesReservationExist === true;
  
  // Log para debug
  console.log('🔄 actualizarEstadoBotonCrearReserva - doesReservationExist:', doesReservationExist, 'extractionState:', extractionState);
  
  // Actualizar el texto del botón según el estado (usar doesReservationExist para crear/editar)
  actualizarTextoBotonReserva(doesReservationExist);
  
  if (esValido) {
    boton.disabled = false;
    boton.style.opacity = "1";
    boton.style.cursor = "pointer";
  } else {
    boton.disabled = true;
    boton.style.opacity = "0.5";
    boton.style.cursor = "not-allowed";
  }
}

/**
 * Deshabilita todos los campos del formulario
 */
function deshabilitarFormularios() {
  // Deshabilitar campos de pasajeros
  const container = document.getElementById("pasajerosContainer");
  if (container) {
    const inputs = container.querySelectorAll("input, select");
    inputs.forEach(input => {
      input.disabled = true;
      input.style.backgroundColor = "#f3f4f6";
      input.style.cursor = "not-allowed";
    });
    
    // Deshabilitar botones de eliminar pasajero
    const botonesEliminar = container.querySelectorAll(".btn-eliminar-pasajero");
    botonesEliminar.forEach(btn => {
      btn.disabled = true;
      btn.style.opacity = "0.3";
      btn.style.cursor = "not-allowed";
      btn.style.pointerEvents = "none";
    });
  }
  
  // Deshabilitar campos de reserva
  const camposReserva = ['tipoReserva', 'estadoReserva', 'fechaViaje', 'vendedor', 'cliente', 'moneda'];
  camposReserva.forEach(campoId => {
    const campo = document.getElementById(campoId);
    if (campo) {
      campo.disabled = true;
      campo.style.backgroundColor = "#f3f4f6";
      campo.style.cursor = "not-allowed";
    }
  });
  
  // Deshabilitar botón de agregar pasajero
  const btnAgregar = document.getElementById("agregarPasajero");
  if (btnAgregar) {
    btnAgregar.disabled = true;
    btnAgregar.style.opacity = "0.5";
    btnAgregar.style.cursor = "not-allowed";
  }
  
  // Deshabilitar botón de guardar
  const btnGuardar = document.getElementById("guardar");
  if (btnGuardar) {
    btnGuardar.disabled = true;
    btnGuardar.style.opacity = "0.5";
    btnGuardar.style.cursor = "not-allowed";
  }
}

/**
 * Resetea la aplicación al estado inicial
 */
function resetearAplicacion() {
  // Ocultar resultados
  const resultsDiv = document.getElementById("results");
  if (resultsDiv) {
    resultsDiv.style.display = "none";
  }
  
  // Mostrar botón de extraer
  const runButton = document.getElementById("run");
  if (runButton) {
    runButton.style.display = "block";
  }
  
  // Ocultar botón de re-extracción
  const reextractButton = document.getElementById("reextract");
  if (reextractButton) {
    reextractButton.style.display = "none";
  }
  
  // Limpiar contenedores y asegurar que estén habilitados
  const pasajerosContainer = document.getElementById("pasajerosContainer");
  if (pasajerosContainer) {
    pasajerosContainer.innerHTML = "";
    // Asegurar que el contenedor esté visible
    const resultsDiv = document.getElementById("results");
    if (resultsDiv) {
      // No forzar display aquí, se mostrará cuando se extraiga
    }
  }
  
  // Resetear estado de extracción
  extractionState = {
    didExtractionExist: false,
    doesReservationExist: false,
    reservationCode: null,
    originData: null,
    createdReservationCode: null
  };
  
  // Restaurar HTML original de datosReservaSection
  const datosReservaSection = document.getElementById("datosReservaSection");
  if (datosReservaSection && datosReservaSectionOriginalHTML) {
    datosReservaSection.innerHTML = datosReservaSectionOriginalHTML;
    // Repoblar los selects después de restaurar
    setTimeout(() => {
      poblarSelectReserva();
      // Asegurar que los campos estén habilitados
      const camposReserva = ['tipoReserva', 'estadoReserva', 'fechaViaje', 'vendedor', 'cliente', 'moneda'];
      camposReserva.forEach(campoId => {
        const campo = document.getElementById(campoId);
        if (campo) {
          campo.disabled = false;
          campo.style.backgroundColor = "";
          campo.style.cursor = "";
          campo.value = "";
        }
      });
    }, 100);
  } else {
    // Si no hay HTML original guardado, limpiar campos manualmente
    const camposReserva = ['tipoReserva', 'estadoReserva', 'fechaViaje', 'vendedor', 'cliente', 'moneda'];
    camposReserva.forEach(campoId => {
      const campo = document.getElementById(campoId);
      if (campo) {
        campo.disabled = false;
        campo.style.backgroundColor = "";
        campo.style.cursor = "";
        campo.value = "";
      }
    });
  }
  
  // Ocultar secciones de hotel y servicios
  const hotelSection = document.getElementById("hotelSection");
  if (hotelSection) {
    hotelSection.style.display = "none";
  }
  
  const serviciosSection = document.getElementById("serviciosSection");
  if (serviciosSection) {
    serviciosSection.style.display = "none";
  }
  
  const vuelosSection = document.getElementById("vuelosSection");
  if (vuelosSection) {
    vuelosSection.style.display = "none";
  }
  
  // Ocultar botón de agregar vuelo
  const agregarVueloButton = document.getElementById("agregarVuelo");
  if (agregarVueloButton) {
    agregarVueloButton.style.display = "none";
  }
  
  // Ocultar botón de crear otra reserva si existe
  const botonCrearOtra = document.getElementById("crearOtraReserva");
  if (botonCrearOtra) {
    botonCrearOtra.remove();
  }
}

/**
 * Convierte los formularios a modo lectura (texto plano)
 * @param {string} reservationCode - Código de la reserva creada/editada
 */
function convertirAModoLectura(reservationCode = null) {
  // Convertir pasajeros a modo lectura
  const container = document.getElementById("pasajerosContainer");
  if (container) {
    const pasajeros = container.querySelectorAll(".pasajero-acordeon");
    
    pasajeros.forEach((pasajeroDiv) => {
      const numero = pasajeroDiv.dataset.numeroPasajero;
      const content = pasajeroDiv.querySelector(".pasajero-content");
      
      if (content) {
        // Obtener valores actuales
        const datos = {
          tipoPasajero: content.querySelector(`#tipoPasajero_${numero}`)?.value || "",
          nombre: content.querySelector(`#nombre_${numero}`)?.value || "",
          apellido: content.querySelector(`#apellido_${numero}`)?.value || "",
          dni: content.querySelector(`#dni_${numero}`)?.value || "",
          fechaNacimiento: content.querySelector(`#fechaNacimiento_${numero}`)?.value || "",
          cuil: content.querySelector(`#cuil_${numero}`)?.value || "",
          tipoDoc: content.querySelector(`#tipoDoc_${numero}`)?.value || "",
          sexo: content.querySelector(`#sexo_${numero}`)?.value || "",
          nacionalidad: content.querySelector(`#nacionalidad_${numero}`)?.value || "",
          direccion: content.querySelector(`#direccion_${numero}`)?.value || "",
          telefono: content.querySelector(`#telefono_${numero}`)?.value || ""
        };
        
        // Crear HTML en modo lectura
        content.innerHTML = `
          <div class="campo-lectura">
            <label>Tipo de Pasajero:</label>
            <p>${datos.tipoPasajero || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Nombre:</label>
            <p>${datos.nombre || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Apellido:</label>
            <p>${datos.apellido || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>DNI:</label>
            <p>${datos.dni || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Fecha de Nacimiento:</label>
            <p>${datos.fechaNacimiento || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>CUIL:</label>
            <p>${datos.cuil || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Tipo de Documento:</label>
            <p>${datos.tipoDoc || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Sexo:</label>
            <p>${datos.sexo || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Nacionalidad:</label>
            <p>${datos.nacionalidad || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Dirección:</label>
            <p>${datos.direccion || '-'}</p>
          </div>
          <div class="campo-lectura">
            <label>Número de Teléfono:</label>
            <p>${datos.telefono || '-'}</p>
          </div>
        `;
      }
      
      // Ocultar botón de eliminar
      const btnEliminar = pasajeroDiv.querySelector(".btn-eliminar-pasajero");
      if (btnEliminar) {
        btnEliminar.style.display = "none";
      }
    });
  }
  
  // Convertir datos de reserva a modo lectura
  const datosReservaSection = document.getElementById("datosReservaSection");
  if (datosReservaSection) {
    // Usar el código de reserva pasado como parámetro o el guardado en el estado
    const codigoReserva = reservationCode || extractionState.createdReservationCode || null;
    
    const datosReserva = {
      tipoReserva: document.getElementById("tipoReserva")?.value || "",
      estadoReserva: document.getElementById("estadoReserva")?.value || "",
      fechaViaje: document.getElementById("fechaViaje")?.value || "",
      vendedor: document.getElementById("vendedor")?.value || "",
      cliente: document.getElementById("cliente")?.value || ""
    };
    
    datosReservaSection.innerHTML = `
      <h3>Datos de la Reserva</h3>
      ${codigoReserva ? `
      <div class="campo-lectura">
        <label>Código de Reserva:</label>
        <p><strong>${codigoReserva}</strong></p>
      </div>
      ` : ''}
      <div class="campo-lectura">
        <label>Tipo de Reserva:</label>
        <p>${datosReserva.tipoReserva || '-'}</p>
      </div>
      <div class="campo-lectura">
        <label>Estado:</label>
        <p>${datosReserva.estadoReserva || '-'}</p>
      </div>
      <div class="campo-lectura">
        <label>Fecha de Viaje:</label>
        <p>${datosReserva.fechaViaje || '-'}</p>
      </div>
      <div class="campo-lectura">
        <label>Vendedor:</label>
        <p>${datosReserva.vendedor || '-'}</p>
      </div>
      <div class="campo-lectura">
        <label>Cliente:</label>
        <p>${datosReserva.cliente || '-'}</p>
      </div>
    `;
    
    // Agregar botón "Crear otra reserva" dentro de la misma sección
    if (!document.getElementById("crearOtraReserva")) {
      const botonCrearOtra = document.createElement("button");
      botonCrearOtra.id = "crearOtraReserva";
      botonCrearOtra.className = "ms-Button";
      botonCrearOtra.style.cssText = "margin-top: 20px; padding: 12px 24px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;";
      botonCrearOtra.innerHTML = '<span class="ms-Button-label">➕ Crear otra reserva</span>';
      botonCrearOtra.onclick = function() {
        resetearAplicacion();
      };
      
      // Insertar después de la sección de datos de reserva
      datosReservaSection.appendChild(botonCrearOtra);
    }
  }
  
  // Ocultar botón de agregar pasajero
  const btnAgregar = document.getElementById("agregarPasajero");
  if (btnAgregar) {
    btnAgregar.style.display = "none";
  }
  
  // Ocultar botón de guardar
  const btnGuardar = document.getElementById("guardar");
  if (btnGuardar) {
    btnGuardar.style.display = "none";
  }
  
  // Ocultar botón de crear reserva
  const btnCrearReserva = document.getElementById("crearReserva");
  if (btnCrearReserva) {
    btnCrearReserva.style.display = "none";
  }
}

/**
 * Compara dos objetos de datos para detectar cambios en pasajeros, servicios y hotel
 * @param {Object} originData - Data original de la extracción
 * @param {Array} nuevosPasajeros - Pasajeros actuales del formulario
 * @param {Array} nuevosServicios - Servicios actuales del formulario
 * @param {Object} nuevoHotel - Hotel actual del formulario
 * @returns {Object} Objeto con información sobre los cambios detectados
 */
function compararDatos(originData, nuevosPasajeros, nuevosServicios, nuevoHotel) {
  const cambios = {
    pasajeros: false,
    servicios: false,
    hotel: false,
    tieneCambios: false
  };
  
  if (!originData) {
    return cambios;
  }
  
  // Comparar pasajeros
  const pasajerosOriginales = originData.passengers || [];
  if (pasajerosOriginales.length !== nuevosPasajeros.length) {
    cambios.pasajeros = true;
    cambios.tieneCambios = true;
  } else {
    // Comparar cada pasajero
    for (let i = 0; i < nuevosPasajeros.length; i++) {
      const nuevo = nuevosPasajeros[i];
      const original = pasajerosOriginales[i];
      
      if (!original || 
          nuevo.nombre !== (original.firstName || '') ||
          nuevo.apellido !== (original.lastName || '') ||
          nuevo.dni !== (original.documentNumber || '') ||
          nuevo.fechaNacimiento !== (original.dateOfBirth || '')) {
        cambios.pasajeros = true;
        cambios.tieneCambios = true;
        break;
      }
    }
  }
  
  // Comparar servicios
  const serviciosOriginales = originData.services || [];
  if (serviciosOriginales.length !== nuevosServicios.length) {
    cambios.servicios = true;
    cambios.tieneCambios = true;
  } else {
    // Comparar cada servicio
    for (let i = 0; i < nuevosServicios.length; i++) {
      const nuevo = nuevosServicios[i];
      const original = serviciosOriginales[i];
      
      if (!original ||
          nuevo.destino !== (original.destino || '') ||
          nuevo.in !== (original.in || '') ||
          nuevo.out !== (original.out || '') ||
          nuevo.servicio !== (original.servicio || '')) {
        cambios.servicios = true;
        cambios.tieneCambios = true;
        break;
      }
    }
  }
  
  // Comparar hotel
  const hotelOriginal = originData.hotel || null;
  if ((hotelOriginal === null && nuevoHotel !== null) ||
      (hotelOriginal !== null && nuevoHotel === null)) {
    cambios.hotel = true;
    cambios.tieneCambios = true;
  } else if (hotelOriginal && nuevoHotel) {
    if (hotelOriginal.nombre_hotel !== nuevoHotel.nombre_hotel ||
        hotelOriginal.Ciudad !== nuevoHotel.Ciudad ||
        hotelOriginal.in !== nuevoHotel.in ||
        hotelOriginal.out !== nuevoHotel.out) {
      cambios.hotel = true;
      cambios.tieneCambios = true;
    }
  }
  
  return cambios;
}

/**
 * Muestra un modal de confirmación cuando hay cambios en pasajeros, servicios o hotel
 * @param {Object} cambios - Objeto con información sobre los cambios detectados
 * @returns {Promise<boolean>} true si el usuario confirma, false si cancela
 */
function mostrarModalConfirmacion(cambios) {
  return new Promise((resolve) => {
    // Crear overlay
    const overlay = document.createElement('div');
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      z-index: 10001;
      display: flex;
      justify-content: center;
      align-items: center;
    `;
    
    // Crear modal
    const modal = document.createElement('div');
    modal.style.cssText = `
      background: white;
      padding: 24px;
      border-radius: 8px;
      max-width: 500px;
      width: 90%;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    `;
    
    // Lista de cambios
    const cambiosLista = [];
    if (cambios.pasajeros) cambiosLista.push('Pasajeros');
    if (cambios.servicios) cambiosLista.push('Servicios');
    if (cambios.hotel) cambiosLista.push('Hotel');
    
    modal.innerHTML = `
      <h3 style="margin-top: 0; margin-bottom: 16px; color: #d32f2f;">⚠️ Advertencia</h3>
      <p style="margin-bottom: 16px; line-height: 1.5;">
        Se detectaron cambios en los siguientes elementos que serán <strong>reemplazados</strong>:
      </p>
      <ul style="margin-bottom: 20px; padding-left: 20px;">
        ${cambiosLista.map(item => `<li style="margin-bottom: 8px;">${item}</li>`).join('')}
      </ul>
      <p style="margin-bottom: 20px; color: #666; font-size: 14px;">
        ¿Desea continuar con la edición?
      </p>
      <div style="display: flex; gap: 12px; justify-content: flex-end;">
        <button id="modalCancelar" style="
          padding: 10px 20px;
          border: 1px solid #ccc;
          background: white;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        ">Cancelar</button>
        <button id="modalContinuar" style="
          padding: 10px 20px;
          border: none;
          background: #0078d4;
          color: white;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        ">Continuar</button>
      </div>
    `;
    
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    
    // Event listeners
    document.getElementById('modalCancelar').onclick = () => {
      document.body.removeChild(overlay);
      resolve(false);
    };
    
    document.getElementById('modalContinuar').onclick = () => {
      document.body.removeChild(overlay);
      resolve(true);
    };
    
    // Cerrar al hacer click fuera del modal
    overlay.onclick = (e) => {
      if (e.target === overlay) {
        document.body.removeChild(overlay);
        resolve(false);
      }
    };
  });
}

/**
 * Ejecuta la creación de reserva en iTraffic usando RPA
 */
async function ejecutarCrearReserva() {
  try {
    // Obtener datos de todos los pasajeros
    const container = document.getElementById("pasajerosContainer");
    if (!container) {
      mostrarMensaje("No se encontró el contenedor de pasajeros", "error");
      return;
    }
    
    const pasajeros = container.querySelectorAll(".pasajero-acordeon");
    if (pasajeros.length === 0) {
      mostrarMensaje("No hay pasajeros para crear la reserva. Por favor, extraiga datos primero.", "info");
      return;
    }
    
    // Recopilar datos de todos los pasajeros
    const todosPasajeros = [];
    pasajeros.forEach((pasajeroDiv, index) => {
      const numero = pasajeroDiv.dataset.numeroPasajero;
      const content = pasajeroDiv.querySelector(".pasajero-content");
      
      if (content) {
        const datos = {
          numeroPasajero: index + 1,
          tipoPasajero: content.querySelector(`#tipoPasajero_${numero}`)?.value || "",
          nombre: content.querySelector(`#nombre_${numero}`)?.value || "",
          apellido: content.querySelector(`#apellido_${numero}`)?.value || "",
          dni: content.querySelector(`#dni_${numero}`)?.value || "",
          fechaNacimiento: content.querySelector(`#fechaNacimiento_${numero}`)?.value || "",
          cuil: content.querySelector(`#cuil_${numero}`)?.value || "",
          tipoDoc: content.querySelector(`#tipoDoc_${numero}`)?.value || "",
          sexo: content.querySelector(`#sexo_${numero}`)?.value || "",
          nacionalidad: content.querySelector(`#nacionalidad_${numero}`)?.value || "",
          direccion: content.querySelector(`#direccion_${numero}`)?.value || "",
          telefono: content.querySelector(`#telefono_${numero}`)?.value || ""
        };
        
        todosPasajeros.push(datos);
      }
    });
    
    // Capturar datos de la reserva
    const datosReserva = {
      tipoReserva: document.getElementById("tipoReserva")?.value || "",
      estadoReserva: document.getElementById("estadoReserva")?.value || "",
      fechaViaje: document.getElementById("fechaViaje")?.value || "",
      vendedor: document.getElementById("vendedor")?.value || "",
      cliente: document.getElementById("cliente")?.value || "",
      codigo: document.getElementById("codigo")?.value || "",
      reservationDate: document.getElementById("fechaReserva")?.value || "",
      tourEndDate: document.getElementById("fechaFinTour")?.value || "",
      dueDate: document.getElementById("fechaVencimiento")?.value || "",
      contact: document.getElementById("contacto")?.value || "",
      contactEmail: document.getElementById("contactEmail")?.value || "",
      contactPhone: document.getElementById("contactPhone")?.value || "",
      currency: document.getElementById("moneda")?.value || "",
      exchangeRate: parseFloat(document.getElementById("tipoCambio")?.value) || 0,
      commission: parseFloat(document.getElementById("comision")?.value) || 0,
      netAmount: parseFloat(document.getElementById("montoNeto")?.value) || 0,
      grossAmount: parseFloat(document.getElementById("montoBruto")?.value) || 0,
      tripName: document.getElementById("nombreViaje")?.value || "",
      productCode: document.getElementById("codigoProducto")?.value || "",
      adults: parseInt(document.getElementById("adultos")?.value) || 0,
      children: parseInt(document.getElementById("ninos")?.value) || 0,
      infants: parseInt(document.getElementById("infantes")?.value) || 0,
      provider: document.getElementById("proveedor")?.value || "",
      reservationCode: document.getElementById("codigoReserva")?.value || extractionState.reservationCode || "",
      hotel: (() => {
        // Verificar primero si la sección está visible
        const hotelSection = document.getElementById("hotelSection");
        if (!hotelSection || hotelSection.style.display === "none") {
          return null;
        }
        
        const nombreHotel = document.getElementById("hotel_nombre")?.value || "";
        const tipoHabitacion = document.getElementById("hotel_tipo_habitacion")?.value || "";
        const ciudad = document.getElementById("hotel_ciudad")?.value || "";
        const categoria = document.getElementById("hotel_categoria")?.value || null;
        const prioridad = document.getElementById("hotel_prioridad")?.value || null;
        const hotelIn = document.getElementById("hotel_in")?.value || "";
        const hotelOut = document.getElementById("hotel_out")?.value || "";
        
        // Si todos los campos obligatorios están vacíos, retornar null
        if (!nombreHotel && !tipoHabitacion && !ciudad && !hotelIn && !hotelOut) {
          return null;
        }
        
        return {
          nombre_hotel: nombreHotel,
          tipo_habitacion: tipoHabitacion,
          Ciudad: ciudad,
          Categoria: categoria || null,
          Prioridad: prioridad || null,
          in: hotelIn,
          out: hotelOut
        };
      })(),
      checkIn: document.getElementById("checkIn")?.value || "",
      checkOut: document.getElementById("checkOut")?.value || "",
      estadoDeuda: document.getElementById("estadoDeuda")?.value || ""
    };
    
    // Agregar conversationId e itemId del email actual
    const item = Office.context.mailbox.item;
    datosReserva.conversationId = item.conversationId || null;
    datosReserva.itemId = item.itemId || null; // ID del email para Graph API
    
    // Capturar servicios
    const servicios = [];
    let serviciosSection = document.getElementById("serviciosSection");
    const serviciosContainer = document.getElementById("serviciosContainer");
    // Solo capturar servicios si la sección está visible y hay elementos
    if (serviciosSection && serviciosSection.style.display !== "none" && serviciosContainer) {
      const servicioItems = serviciosContainer.querySelectorAll(".servicio-item");
      if (servicioItems.length > 0) {
        servicioItems.forEach((item, index) => {
          servicios.push({
            destino: document.getElementById(`servicio_destino_${index}`)?.value || "",
            in: document.getElementById(`servicio_in_${index}`)?.value || "",
            out: document.getElementById(`servicio_out_${index}`)?.value || "",
            nts: parseInt(document.getElementById(`servicio_nts_${index}`)?.value) || 0,
            basePax: parseInt(document.getElementById(`servicio_basePax_${index}`)?.value) || 0,
            servicio: document.getElementById(`servicio_servicio_${index}`)?.value || "",
            descripcion: document.getElementById(`servicio_descripcion_${index}`)?.value || "",
            estado: document.getElementById(`servicio_estado_${index}`)?.value || "",
            categoria: document.getElementById(`servicio_categoria_${index}`)?.value || "",
            prioridad: document.getElementById(`servicio_prioridad_${index}`)?.value || ""
          });
        });
      }
    }
    // Siempre asignar array vacío si no hay servicios
    datosReserva.services = servicios;
    
    // Capturar vuelos
    const vuelos = [];
    const vuelosSection = document.getElementById("vuelosSection");
    const vuelosContainer = document.getElementById("vuelosContainer");
    // Solo capturar vuelos si la sección está visible y hay elementos
    if (vuelosSection && vuelosSection.style.display !== "none" && vuelosContainer) {
      const vueloItems = vuelosContainer.querySelectorAll(".vuelo-item");
      if (vueloItems.length > 0) {
        vueloItems.forEach((item, index) => {
          vuelos.push({
            airline: document.getElementById(`vuelo_airline_${index}`)?.value || "",
            flightNumber: document.getElementById(`vuelo_flightNumber_${index}`)?.value || "",
            origin: document.getElementById(`vuelo_origin_${index}`)?.value || "",
            destination: document.getElementById(`vuelo_destination_${index}`)?.value || "",
            departureDate: document.getElementById(`vuelo_departureDate_${index}`)?.value || "",
            departureTime: document.getElementById(`vuelo_departureTime_${index}`)?.value || "",
            arrivalDate: document.getElementById(`vuelo_arrivalDate_${index}`)?.value || "",
            arrivalTime: document.getElementById(`vuelo_arrivalTime_${index}`)?.value || ""
          });
        });
      }
    }
    // Siempre asignar array vacío si no hay vuelos
    datosReserva.flights = vuelos;
    
    // VALIDAR CAMPOS OBLIGATORIOS
    let camposFaltantes = [];
    
    // Validar pasajeros (solo nombre y apellido son obligatorios)
    todosPasajeros.forEach((pasajero, index) => {
      if (!pasajero.nombre || pasajero.nombre.trim() === "") {
        camposFaltantes.push(`Pasajero ${index + 1}: Nombre`);
      }
      if (!pasajero.apellido || pasajero.apellido.trim() === "") {
        camposFaltantes.push(`Pasajero ${index + 1}: Apellido`);
      }
    });
    
    // Validar datos de reserva (todos obligatorios)
    if (!datosReserva.tipoReserva || datosReserva.tipoReserva.trim() === "") {
      camposFaltantes.push("Tipo de Reserva");
    }
    if (!datosReserva.estadoReserva || datosReserva.estadoReserva.trim() === "") {
      camposFaltantes.push("Estado");
    }
    if (!datosReserva.fechaViaje || datosReserva.fechaViaje.trim() === "") {
      camposFaltantes.push("Fecha de Viaje");
    }
    if (!datosReserva.vendedor || datosReserva.vendedor.trim() === "") {
      camposFaltantes.push("Vendedor");
    }
    if (!datosReserva.cliente || datosReserva.cliente.trim() === "") {
      camposFaltantes.push("Cliente");
    }
    
    // Validar campos obligatorios del hotel (solo si la sección está visible y tiene campos required)
    const hotelSection = document.getElementById("hotelSection");
    if (hotelSection && hotelSection.style.display !== "none") {
      const hotelNombre = document.getElementById("hotel_nombre");
      const hotelTipoHabitacion = document.getElementById("hotel_tipo_habitacion");
      const hotelCiudad = document.getElementById("hotel_ciudad");
      const hotelIn = document.getElementById("hotel_in");
      const hotelOut = document.getElementById("hotel_out");
      
      // Solo validar si los campos son required
      if (hotelNombre?.required && (!datosReserva.hotel?.nombre_hotel || datosReserva.hotel.nombre_hotel.trim() === "")) {
        camposFaltantes.push("Hotel: Nombre");
      }
      if (hotelTipoHabitacion?.required && (!datosReserva.hotel?.tipo_habitacion || datosReserva.hotel.tipo_habitacion.trim() === "")) {
        camposFaltantes.push("Hotel: Tipo de Habitación");
      }
      if (hotelCiudad?.required && (!datosReserva.hotel?.Ciudad || datosReserva.hotel.Ciudad.trim() === "")) {
        camposFaltantes.push("Hotel: Ciudad");
      }
      if (hotelIn?.required && (!datosReserva.hotel?.in || datosReserva.hotel.in.trim() === "")) {
        camposFaltantes.push("Hotel: Fecha Entrada");
      }
      if (hotelOut?.required && (!datosReserva.hotel?.out || datosReserva.hotel.out.trim() === "")) {
        camposFaltantes.push("Hotel: Fecha Salida");
      }
      
      // Validar que las fechas de entrada y salida tengan al menos un día de diferencia
      if (datosReserva.hotel?.in && datosReserva.hotel?.out) {
        const fechaIn = new Date(datosReserva.hotel.in);
        const fechaOut = new Date(datosReserva.hotel.out);
        
        // Verificar que la fecha de salida sea posterior a la de entrada
        if (fechaOut <= fechaIn) {
          mostrarMensaje("Las fechas del hotel no son válidas. La fecha de salida debe ser al menos un día posterior a la fecha de entrada.", "error");
          return; // NO enviar al RPA
        }
        
        // Calcular la diferencia en días
        const diferenciaDias = Math.floor((fechaOut - fechaIn) / (1000 * 60 * 60 * 24));
        
        if (diferenciaDias < 1) {
          mostrarMensaje("Las fechas del hotel deben tener al menos un día de diferencia. La fecha de salida debe ser posterior a la fecha de entrada.", "error");
          return; // NO enviar al RPA
        }
      }
    }
    
    // Validar campos obligatorios de servicios (solo si la sección está visible)
    // Reutilizar la variable serviciosSection ya declarada arriba
    if (serviciosSection && serviciosSection.style.display !== "none") {
      if (!datosReserva.services || datosReserva.services.length === 0) {
        // Solo requerir servicios si hay algún campo required en la sección
        const serviciosContainer = document.getElementById("serviciosContainer");
        if (serviciosContainer) {
          const primerServicio = serviciosContainer.querySelector(".servicio-item");
          if (primerServicio) {
            const tieneRequired = primerServicio.querySelector("[required]");
            if (tieneRequired) {
              camposFaltantes.push("Servicios: Debe haber al menos un servicio");
            }
          }
        }
      } else {
        datosReserva.services.forEach((servicio, index) => {
          const destino = document.getElementById(`servicio_destino_${index}`);
          const servicioIn = document.getElementById(`servicio_in_${index}`);
          const servicioOut = document.getElementById(`servicio_out_${index}`);
          const nts = document.getElementById(`servicio_nts_${index}`);
          const basePax = document.getElementById(`servicio_basePax_${index}`);
          const servicioField = document.getElementById(`servicio_servicio_${index}`);
          const estado = document.getElementById(`servicio_estado_${index}`);
          
          // Solo validar si los campos son required
          if (destino?.required && (!servicio.destino || servicio.destino.trim() === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Destino`);
          }
          if (servicioIn?.required && (!servicio.in || servicio.in.trim() === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Fecha Entrada`);
          }
          if (servicioOut?.required && (!servicio.out || servicio.out.trim() === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Fecha Salida`);
          }
          if (nts?.required && (nts.value === null || nts.value === undefined || nts.value === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Noches`);
          }
          if (basePax?.required && (!servicio.basePax || servicio.basePax === 0)) {
            camposFaltantes.push(`Servicio ${index + 1}: Base Pax`);
          }
          if (servicioField?.required && (!servicio.servicio || servicio.servicio.trim() === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Servicio`);
          }
          if (estado?.required && (!servicio.estado || servicio.estado.trim() === "")) {
            camposFaltantes.push(`Servicio ${index + 1}: Estado`);
          }
        });
      }
    }
    
    // Si hay campos faltantes, mostrar error y no enviar
    if (camposFaltantes.length > 0) {
      mostrarMensaje(
        `Por favor completa los siguientes campos obligatorios: ${camposFaltantes.join(', ')}`,
        "error"
      );
      return; // NO enviar al RPA
    }
    
    // Obtener el estado de si la reserva existe (convertir explícitamente a booleano)
    const doesReservationExist = extractionState.doesReservationExist === true;
    
    // Log para debug
    console.log('🚀 ejecutarCrearReserva - doesReservationExist:', doesReservationExist, 'extractionState:', extractionState);
    
    // Si es edición, validar cambios antes de continuar (usar doesReservationExist)
    if (doesReservationExist && extractionState.originData) {
      // Capturar servicios y hotel para comparación
      const servicios = datosReserva.services || [];
      const hotel = datosReserva.hotel || null;
      
      // Comparar datos
      const cambios = compararDatos(extractionState.originData, todosPasajeros, servicios, hotel);
      
      // Si hay cambios, mostrar modal de confirmación
      if (cambios.tieneCambios) {
        const confirmar = await mostrarModalConfirmacion(cambios);
        if (!confirmar) {
          // El usuario canceló, no hacer nada
          return;
        }
      }
    }
    
    // Mostrar mensaje de procesamiento según si es crear o editar (usar doesReservationExist)
    const mensajeProcesamiento = doesReservationExist 
      ? "Editando reserva en iTraffic... Por favor espere." 
      : "Creando reserva en iTraffic... Por favor espere.";
    mostrarMensaje(mensajeProcesamiento, "info");
    
    // DESHABILITAR TODOS LOS CAMPOS
    deshabilitarFormularios();
    
    // Deshabilitar botón mientras se procesa
    const botonCrearReserva = document.getElementById("crearReserva");
    if (botonCrearReserva) {
      botonCrearReserva.disabled = true;
      botonCrearReserva.style.opacity = "0.6";
      botonCrearReserva.querySelector('.ms-Button-label').textContent = "⏳ Procesando...";
    }
    
    // Llamar al servicio RPA con los datos de pasajeros y reserva, pasando el estado y la data original si es edición
    const originData = doesReservationExist ? extractionState.originData : null;
    const resultado = await crearReservaEnITraffic(todosPasajeros, datosReserva, doesReservationExist, originData);
    
    // Obtener reservationCode de la respuesta
    const reservationCode = resultado.data?.reservationCode || resultado.reservationCode || null;
    
    // Guardar el código de reserva en el estado
    if (reservationCode) {
      extractionState.createdReservationCode = reservationCode;
    }
    
    // Mostrar mensaje de éxito según si es crear o editar, incluyendo el código de reserva si existe (usar doesReservationExist)
    let mensajeExito = doesReservationExist 
      ? "¡Reserva editada exitosamente en iTraffic!" 
      : "¡Reserva creada exitosamente en iTraffic!";
    
    if (reservationCode) {
      mensajeExito += ` Código de la reserva: ${reservationCode}`;
    }
    
    mostrarMensaje(mensajeExito, "success");
    
    // CONVERTIR A MODO LECTURA (pasar el código de reserva)
    convertirAModoLectura(reservationCode);
    
  } catch (error) {
    const doesReservationExist = extractionState.doesReservationExist === true;
    const mensajeError = doesReservationExist 
      ? "Error al editar reserva: " + error.message 
      : "Error al crear reserva: " + error.message;
    mostrarMensaje(mensajeError, "error");
    // En caso de error, rehabilitar los formularios
    habilitarFormularios();
  }
}

/**
 * Habilita todos los campos del formulario (en caso de error)
 */
function habilitarFormularios() {
  // Habilitar campos de pasajeros
  const container = document.getElementById("pasajerosContainer");
  if (container) {
    const inputs = container.querySelectorAll("input, select");
    inputs.forEach(input => {
      input.disabled = false;
      input.style.backgroundColor = "#ffffff";
      input.style.cursor = "text";
    });
    
    // Habilitar botones de eliminar pasajero
    const botonesEliminar = container.querySelectorAll(".btn-eliminar-pasajero");
    botonesEliminar.forEach(btn => {
      btn.disabled = false;
      btn.style.opacity = "1";
      btn.style.cursor = "pointer";
      btn.style.pointerEvents = "auto";
    });
  }
  
  // Habilitar campos de reserva
  const camposReserva = ['tipoReserva', 'estadoReserva', 'fechaViaje', 'vendedor', 'cliente', 'moneda'];
  camposReserva.forEach(campoId => {
    const campo = document.getElementById(campoId);
    if (campo) {
      campo.disabled = false;
      campo.style.backgroundColor = "#ffffff";
      campo.style.cursor = "pointer";
    }
  });
  
  // Habilitar botón de agregar pasajero
  const btnAgregar = document.getElementById("agregarPasajero");
  if (btnAgregar) {
    btnAgregar.disabled = false;
    btnAgregar.style.opacity = "1";
    btnAgregar.style.cursor = "pointer";
  }
  
  // Habilitar botón de guardar
  const btnGuardar = document.getElementById("guardar");
  if (btnGuardar) {
    btnGuardar.disabled = false;
    btnGuardar.style.opacity = "1";
    btnGuardar.style.cursor = "pointer";
  }
  
  // Habilitar botón de crear reserva
  const botonCrearReserva = document.getElementById("crearReserva");
  if (botonCrearReserva) {
    const doesReservationExist = extractionState.doesReservationExist === true;
    botonCrearReserva.disabled = false;
    botonCrearReserva.style.opacity = "1";
    actualizarTextoBotonReserva(doesReservationExist);
  }
}

export { mostrarMensaje, run, crearFormulariosPasajeros, crearFormularioPasajero, guardarDatos, llenarDatosPasajeros, eliminarPasajero, renumerarPasajeros, agregarNuevoPasajero, ejecutarCrearReserva };
