const express = require("express");

let configViewEngine = (app)=>{
    app.set('view engine', 'ejs');
    app.use(express.static('./src/public'));
    app.set('views', './src/views');
}

module.exports = configViewEngine;