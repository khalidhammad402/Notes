import userService from "../services/userService";
import { validationResult } from "express-validator";

let getHomepage = (req, res)=>{
    return res.render('homepage.ejs')
};

let getRegistrationpage = (req, res) => {
    let form = {
        firstName: req.body.first_name,
        lastName: req.body.last_name,
        email: req.body.email
    }
    return res.render('auth/register.ejs', {
        errors: req.flash("errors"),
        form: form
    })
}

let getLoginpage = (req, res) => {
    return res.render('auth/login.ejs')
}

let registerUser = async (req, res) => {
    let form = {
        firstName: req.body.first_name,
        lastName: req.body.last_name,
        email: req.body.email
    }
    let errorsArr = []
    let validationError = validationResult(req);
    if(!validationError.isEmpty()){
        let errors = Object.values(validationError.mapped());
        errors.forEach((item)=>{
            errorsArr.push(item.msg);
        });
        req.flash("errors", errorsArr);
        return res.render("auth/register.ejs", {
            errors: req.flash("errors"),
            form: form
        });
    }
    try{
        let user = {
            firstName: req.body.first_name,
            lastName: req.body.last_name,
            email: req.body.email,
            password: req.body.password,
            confirmPassword: req.body.password_confirmation,
            createdAt: new Date()
        }
        await userService.createNewUser(user);
        return res.redirect("/");
    } catch(e) {
        req.flash("errors", e);
        return res.render("auth/register.ejs", {
            errors: req.flash("errors"),
            form: form
        });
    }
}

module.exports = {
    getHomepage : getHomepage,
    getRegistrationpage: getRegistrationpage,
    getLoginpage: getLoginpage,
    registerUser: registerUser
}