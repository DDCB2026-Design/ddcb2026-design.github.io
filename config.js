// Configuración de la tienda ABAS-TCT. Edita los valores entre comillas y guarda (Commit changes).
window.ABAS_CONFIG = {
  // Enlace CSV de la pestaña "Tienda" publicada en la web desde Google Sheets.
  // Mientras esté vacío, la web muestra el catálogo de respaldo incluido en index.html.
  sheetCsvUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRm7CBY9jbYeAg1V1Xnoyw-DfZl2fh5VJYexY1JpwdQhTEuFFk7ut7TZYZyPXjk0PDoc6ldsJzfy_yY/pub?gid=2082516956&single=true&output=csv",
  // URL de la aplicación web de Apps Script (registro y consulta de pedidos).
  apiUrl: "https://script.google.com/macros/s/AKfycbzQxekZaPnqyYe2CRrFXMpKt2gIm62l3okxUdP540Uu_O3aXN8z9jTuu5tOyjOj5V_g7A/exec",
  franjas: ["Mañana (9 a 12)", "Tarde (12 a 4)", "Noche (4 a 8)"],
  settings: {
    storeName: "ABAS-TCT",
    tagline: "Frutos secos frescos, directo a tu casa",
    whatsapp: "51987609561",        // con 51 adelante
    yapeNumber: "987609561",                 // ej. "987609561"
    yapeName: "Dante Daniel Clavijo Bazalar",                   // titular tal como aparece en Yape
    yapeQR: "img/qr-pago.png",                     // ej. "img/yape-qr.png" (sube la imagen a la carpeta img)
    delivery: { mode: "district", flat: 12, districts: [{name:"Ancón",cost:18}, {name:"Ate",cost:18}, {name:"Barranco",cost:12}, {name:"Bellavista",cost:18}, {name:"Breña",cost:12}, {name:"Callao",cost:18}, {name:"Carabayllo",cost:18}, {name:"Carmen de la Legua-Reynoso",cost:18}, {name:"Cercado de Lima",cost:12}, {name:"Chaclacayo",cost:18}, {name:"Chorrillos",cost:18}, {name:"Cieneguilla",cost:18}, {name:"Comas",cost:18}, {name:"El Agustino",cost:12}, {name:"Independencia",cost:18}, {name:"Jesús María",cost:12}, {name:"La Molina",cost:18}, {name:"La Perla",cost:18}, {name:"La Punta",cost:18}, {name:"La Victoria",cost:12}, {name:"Lince",cost:12}, {name:"Los Olivos",cost:18}, {name:"Lurigancho-Chosica",cost:18}, {name:"Lurín",cost:18}, {name:"Magdalena del Mar",cost:12}, {name:"Mi Perú",cost:18}, {name:"Miraflores",cost:12}, {name:"Pachacámac",cost:18}, {name:"Pucusana",cost:18}, {name:"Pueblo Libre",cost:12}, {name:"Puente Piedra",cost:18}, {name:"Punta Hermosa",cost:18}, {name:"Punta Negra",cost:18}, {name:"Rímac",cost:12}, {name:"San Bartolo",cost:18}, {name:"San Borja",cost:12}, {name:"San Isidro",cost:12}, {name:"San Juan de Lurigancho",cost:18}, {name:"San Juan de Miraflores",cost:18}, {name:"San Luis",cost:12}, {name:"San Martín de Porres",cost:18}, {name:"San Miguel",cost:12}, {name:"Santa Anita",cost:12}, {name:"Santa María del Mar",cost:18}, {name:"Santa Rosa",cost:18}, {name:"Santiago de Surco",cost:18}, {name:"Surquillo",cost:12}, {name:"Ventanilla",cost:18}, {name:"Villa El Salvador",cost:18}, {name:"Villa María del Triunfo",cost:18}] },  // respaldo; los valores reales se leen de la pestaña Ajustes del Sheet
    pickup: false, pickupAddress: "", minOrder: 50,   // mínimo en S/, incluye el delivery
    logo: "img/logo.jpg", heroImg: "img/hero.jpg",
    catOrder: ["Frutos secos","Pasas y deshidratados","Semillas y granos","Especias","Conservas"],
    catImgs: { "Frutos secos":"img/mix.jpg", "Pasas y deshidratados":"img/deshidratados.jpg", "Semillas y granos":"img/chia.jpg", "Especias":"img/especias.jpg", "Conservas":"img/conservas.jpg" }
  }
};
