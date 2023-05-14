import {check} from "express-validator";

let validateRegister = [
    check("email", "Invalid Email!").isEmail().trim(),

    check("password", "Password must be atleast 3 character long!").isLength({min: 3}),

    check("password_confirmation", "Confirm passwords does not match password")
    .custom((value, {req})=>{
        return value === req.body.password;
    })
]

module.exports = {
    validateRegister : validateRegister
}