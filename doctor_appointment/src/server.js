require("dotenv").config();
import express from "express";
import configViewEngine from "./config/viewEngine";
import initWebRoutes from "./routes/web";
import bodyParser from "body-parser"
import connectFlash from "connect-flash"
import cookieParser from "cookie-parser";
import session from "express-session";

const app = express();

//config cookie-parser
app.use(cookieParser('secret'));

//config express session
app.use(session({
    secret: 'secret',
    resave: true,
    saveUninitialized: false,
    cookie: {
        maxAge: 1000 * 60 * 60 * 24        // 1 day
    }
}));

//show flash messages
app.use(connectFlash());

//config body-parser
app.use(bodyParser.urlencoded({extended: true}));

//config view Engine
configViewEngine(app);

//init all web routes
initWebRoutes(app);

let port = process.env.PORT || 3000;

app.listen(port, ()=>{
   console.log(`App is running at the port ${port}`);
});