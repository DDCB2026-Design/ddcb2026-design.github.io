// Configuración de la tienda ABAS-TCT. Edita los valores entre comillas y guarda (Commit changes).
window.ABAS_CONFIG = {
  // Enlace CSV de la pestaña "Tienda" publicada en la web desde Google Sheets.
  // Mientras esté vacío, la web muestra el catálogo de respaldo incluido en index.html.
  sheetCsvUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRm7CBY9jbYeAg1V1Xnoyw-DfZl2fh5VJYexY1JpwdQhTEuFFk7ut7TZYZyPXjk0PDoc6ldsJzfy_yY/pub?gid=2082516956&single=true&output=csv",
  settings: {
    storeName: "ABAS-TCT",
    tagline: "Frutos secos frescos, directo a tu casa",
    whatsapp: "51987609561",        // con 51 adelante
    yapeNumber: "",                 // ej. "987609561"
    yapeName: "",                   // titular tal como aparece en Yape
    yapeQR: "",                     // ej. "img/yape-qr.png" (sube la imagen a la carpeta img)
    delivery: { mode: "flat", flat: 15, districts: [] },  // tarifa por distrito: mode "district", districts: [{name:"Miraflores", cost:12}]
    pickup: false, pickupAddress: "", minOrder: 0,
    logo: "img/logo.jpg", heroImg: "img/hero.jpg",
    catOrder: ["Frutos secos","Pasas y deshidratados","Semillas y granos","Especias","Conservas"],
    catImgs: { "Frutos secos":"img/mix.jpg", "Pasas y deshidratados":"img/deshidratados.jpg", "Semillas y granos":"img/chia.jpg", "Especias":"img/especias.jpg", "Conservas":"img/conservas.jpg" }
  }
};
