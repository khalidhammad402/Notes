import db from "../models"
import bcrypt from "bcryptjs"

let createNewUser = (user) => {
    return new Promise(async (resolve, reject) => {
      try {
        let doesEmailExist = await emailChecUser(user);
        if(doesEmailExist) {
          reject("This email already exists.")
        }
        else {
          let salt = bcrypt.genSaltSync(10);
          let hashPassword = await bcrypt.hashSync(user.password, salt);
          user.password = hashPassword;
          await db.User.create(user);
          resolve("done!");
        }
      } catch (e) {
        reject(e);
      }
    });
};

let emailChecUser = (user) => {
  return new Promise(async (resolve, reject) => {
    try{
      let currentUser = await db.User.findOne({where: {email: user.email}});
      if(currentUser) resolve(true);
      resolve(false);
    }catch (e) {
      reject(e);
    }
  });
}

module.exports = {
  createNewUser: createNewUser
}