const XLSX = require('xlsx');
const moment = require('moment');

// Leer el archivo Excel
const workbook = XLSX.readFile('Libro1.xlsx');

// Seleccionar la primera hoja de trabajo
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Obtener los datos de la hoja
let data = XLSX.utils.sheet_to_json(worksheet);

console.log('Datos originales:', data);

// Convertir campos de fecha al formato deseado y sumar duración
data = data.map(row => {
  if (row.date) {
    let fecha;
    // Verificar si la fecha es un número de serie de Excel
    if (typeof row.date === 'number') {
      const date = XLSX.SSF.parse_date_code(row.date);
      fecha = moment(new Date(date.y, date.m - 1, date.d));
    } else {
      fecha = moment(row.date, 'DD-MM-YYYY', true); // True para validación estricta
    }

    // Comprobar si la fecha es válida
    if (!fecha.isValid()) {
      console.warn(`Fecha no válida: ${row.date}`);
      return row; // Salta el procesamiento si la fecha es inválida
    }

    // Formatear la fecha original
    row.date = fecha.format('DD-MM-YYYY');

    // Sumar la duración en días a la fecha
    if (row.duracion && !isNaN(row.duracion)) {
      const limitDate = fecha.clone().add(parseInt(row.duracion, 10), 'days');
      // Crear un nuevo campo limit_date con la fecha resultante
      row.limit_date = limitDate.format('DD-MM-YYYY');
    } else {
      row.limit_date = fecha.format('DD-MM-YYYY'); // Si no hay duración, solo copia la fecha
    }
  }
  return row;
});

console.log('Datos con fechas convertidas y limit_date agregado:', data);

// Crear un nuevo arreglo que repita cada objeto aumentando el mes basado en el tipo
let newData = [];

data.forEach(row => {
  let fechaOriginal = moment(row.date, 'DD-MM-YYYY');
  let incremento;
  
  // Establecer el incremento basado en el tipo
  switch (row.tipo) {
    case 'mensual':
      incremento = 1;
      break;
    case 'trimestral':
      incremento = 3;
      break;
    case 'cuatrimestral':
      incremento = 4;
      break;
    case 'bimestral':
      incremento = 2;
      break;
    case 'semestral':
      incremento = 6;
      break;
    case 'anual':
      incremento = 12;
      break;
    case 'bianual':
      incremento = 24;
      break;
    case 'semanal':
      incremento = 'semanal'; // Manejo especial para semanal
      break;
    case 'quincenal':
      incremento = 'quincenal'; // Manejo especial para quincenal
      break;
    default:
      console.warn(`Tipo desconocido: ${row.tipo}`);
      return; // Salta el procesamiento si el tipo es desconocido
  }

  // Generar nuevas entradas basado en el tipo
  if (row.tipo === 'semestral') {
    // Semestral debe tener 2 repeticiones, incluso en el siguiente año
    for (let i = 0; i < 2; i++) {
      let fechaConMesIncrementado = fechaOriginal.clone().add(i * incremento, 'months');
      let nuevoRegistro = { ...row }; // Copiar el objeto original

      // Actualizar el campo de fecha en el nuevo objeto
      nuevoRegistro.date = fechaConMesIncrementado.format('DD-MM-YYYY');

      // Actualizar el campo limit_date si es necesario
      if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
        const limitDate = fechaConMesIncrementado.clone().add(parseInt(nuevoRegistro.duracion, 10), 'days');
        nuevoRegistro.limit_date = limitDate.format('DD-MM-YYYY');
      }

      newData.push(nuevoRegistro); // Agregar el nuevo objeto al arreglo
    }
  } else if (incremento === 'semanal') {
    // Semanal: agregar una semana a la fecha original hasta la última semana del año
    let fechaConSemanaIncrementada = fechaOriginal.clone();
    while (fechaConSemanaIncrementada.year() === fechaOriginal.year()) {
      let nuevoRegistro = { ...row }; // Copiar el objeto original
      nuevoRegistro.date = fechaConSemanaIncrementada.format('DD-MM-YYYY');

      // Actualizar el campo limit_date si es necesario
      if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
        const limitDate = fechaConSemanaIncrementada.clone().add(parseInt(nuevoRegistro.duracion, 10), 'days');
        nuevoRegistro.limit_date = limitDate.format('DD-MM-YYYY');
      }

      newData.push(nuevoRegistro); // Agregar el nuevo objeto al arreglo
      fechaConSemanaIncrementada.add(1, 'week'); // Sumar una semana
    }
  } else if (incremento === 'quincenal') {
    // Quincenal: repetir el mismo día cada mes y sumar 15 días a la segunda fecha
    for (let mes = fechaOriginal.month(); mes <= 11; mes++) {
      let fechaConMesIncrementado = fechaOriginal.clone().month(mes);
      // Verificar si la fecha sigue dentro del mismo año
      if (fechaConMesIncrementado.year() === fechaOriginal.year()) {
        let nuevoRegistro = { ...row }; // Copiar el objeto original

        // Actualizar el campo de fecha en el nuevo objeto
        nuevoRegistro.date = fechaConMesIncrementado.format('DD-MM-YYYY');

        // Actualizar el campo limit_date si es necesario
        if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
          const limitDate = fechaConMesIncrementado.clone().add(parseInt(nuevoRegistro.duracion, 10), 'days');
          nuevoRegistro.limit_date = limitDate.format('DD-MM-YYYY');
        }

        newData.push(nuevoRegistro); // Agregar el nuevo objeto al arreglo
        
        // Agregar una segunda entrada con 15 días adicionales
        let fechaCon15DiasAdicionales = fechaConMesIncrementado.clone().add(15, 'days');
        if (fechaCon15DiasAdicionales.year() === fechaOriginal.year()) {
          let nuevoRegistroCon15Dias = { ...row }; // Copiar el objeto original
          nuevoRegistroCon15Dias.date = fechaCon15DiasAdicionales.format('DD-MM-YYYY');
          
          // Actualizar el campo limit_date si es necesario
          if (nuevoRegistroCon15Dias.duracion && !isNaN(nuevoRegistroCon15Dias.duracion)) {
            const limitDate = fechaCon15DiasAdicionales.clone().add(parseInt(nuevoRegistroCon15Dias.duracion, 10), 'days');
            nuevoRegistroCon15Dias.limit_date = limitDate.format('DD-MM-YYYY');
          }
          
          newData.push(nuevoRegistroCon15Dias); // Agregar el nuevo objeto al arreglo
        }
      }
    }
  } else {
    // Para otros tipos
    let fechaFinDelAño = fechaOriginal.clone().endOf('year');
    let fechaConIncremento = fechaOriginal.clone();
    
    while (fechaConIncremento.isSameOrBefore(fechaFinDelAño)) {
      let nuevoRegistro = { ...row }; // Copiar el objeto original

      // Actualizar el campo de fecha en el nuevo objeto
      nuevoRegistro.date = fechaConIncremento.format('DD-MM-YYYY');

      // Actualizar el campo limit_date si es necesario
      if (nuevoRegistro.duracion && !isNaN(nuevoRegistro.duracion)) {
        const limitDate = fechaConIncremento.clone().add(parseInt(nuevoRegistro.duracion, 10), 'days');
        nuevoRegistro.limit_date = limitDate.format('DD-MM-YYYY');
      }

      newData.push(nuevoRegistro); // Agregar el nuevo objeto al arreglo

      // Incrementar el mes o año según el tipo
      if (row.tipo === 'bianual') {
        fechaConIncremento.add(2 * incremento, 'months');
      } else {
        fechaConIncremento.add(incremento, 'months');
      }
    }
    
    // Manejar el caso bianual para agregar una repetición adicional
    if (row.tipo === 'bianual') {
      let fechaConIncrementoAdicional = fechaOriginal.clone().add(incremento, 'months');
      if (fechaConIncrementoAdicional.year() !== fechaOriginal.year()) {
        let nuevoRegistroAdicional = { ...row }; // Copiar el objeto original
        nuevoRegistroAdicional.date = fechaConIncrementoAdicional.format('DD-MM-YYYY');

        // Actualizar el campo limit_date si es necesario
        if (nuevoRegistroAdicional.duracion && !isNaN(nuevoRegistroAdicional.duracion)) {
          const limitDate = fechaConIncrementoAdicional.clone().add(parseInt(nuevoRegistroAdicional.duracion, 10), 'days');
          nuevoRegistroAdicional.limit_date = limitDate.format('DD-MM-YYYY');
        }
        
        newData.push(nuevoRegistroAdicional); // Agregar el nuevo objeto al arreglo
      }
    }
  }
});

// Eliminar los campos duracion y tipo
newData = newData.map(({ duracion, tipo, ...resto }) => resto);

console.log('Nuevo arreglo con meses incrementados y campos eliminados:', newData);

// Convertir los datos de nuevo a una hoja de trabajo
const newWorksheet = XLSX.utils.json_to_sheet(newData, { header: ['infrastructure_id', 'date', 'limit_date', 'description'] });

// Reemplazar la hoja original con la modificada
workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

// Guardar el archivo Excel actualizado
XLSX.writeFile(workbook, 'Libro1.xlsx');
