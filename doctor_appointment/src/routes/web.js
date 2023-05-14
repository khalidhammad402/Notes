import express from 'express'
import homepageController from '../controllers/homepageController';
import auth from "../validation/authValidation"

let router = express.Router();

let initAtllWebRoutes = (app)=>{
    router.get('/', homepageController.getHomepage);
    router.get('/register', homepageController.getRegistrationpage);
    router.post('/register', auth.validateRegister, homepageController.registerUser)
    app.get("/login", homepageController.getLoginpage)
    return app.use('/', router);
}

module.exports = initAtllWebRoutes;