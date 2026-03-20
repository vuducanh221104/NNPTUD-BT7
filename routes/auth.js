var express = require('express');
var router = express.Router();
let userController = require('../controllers/users')
let { RegisterValidator, handleResultValidator } = require('../utils/validatorHandler')
let bcrypt = require('bcrypt')
let jwt = require('jsonwebtoken')
let {checkLogin} = require('../utils/authHandler')
const fs = require('fs');

const privateKey = fs.readFileSync('./private.pem', 'utf8');
/* GET home page. */
router.post('/register', RegisterValidator, handleResultValidator, async function (req, res, next) {
    let newUser = userController.CreateAnUser(
        req.body.username,
        req.body.password,
        req.body.email,
        "69aa8360450df994c1ce6c4c"
    );
    await newUser.save()
    res.send({
        message: "dang ki thanh cong"
    })
});
router.post('/login', async function (req, res, next) {
    let { username, password } = req.body;
    let getUser = await userController.FindByUsername(username);
    if (!getUser) {
        res.status(403).send("tai khoan khong ton tai")
    } else {
        if (getUser.lockTime && getUser.lockTime > Date.now()) {
            res.status(403).send("tai khoan dang bi ban");
            return;
        }
        if (bcrypt.compareSync(password, getUser.password)) {
            await userController.SuccessLogin(getUser);
            let token = jwt.sign({
                id: getUser._id
            }, privateKey,{
                algorithm: 'RS256',
                expiresIn:'30d'
            })
            res.send(token)
        } else {
            await userController.FailLogin(getUser);
            res.status(403).send("thong tin dang nhap khong dung")
        }
    }

});
router.get('/me',checkLogin,function(req,res,next){
    res.send(req.user)
})

router.post('/changepassword', checkLogin, async function(req, res, next){
    let { oldPassword, newPassword } = req.body;
    let user = req.user;

    if (!oldPassword || !newPassword) {
        res.status(400).send("Vui lòng nhập đầy đủ oldPassword và newPassword");
        return;
    }

    if (!bcrypt.compareSync(oldPassword, user.password)) {
        res.status(400).send("Mật khẩu cũ không đúng");
        return;
    }

    // Validate new password
    if (newPassword.length < 6) {
        res.status(400).send("Mật khẩu mới phải có ít nhất 6 ký tự");
        return;
    }
    if (newPassword.length > 20) {
        res.status(400).send("Mật khẩu mới không được quá 20 ký tự");
        return;
    }
    if (!/[A-Za-z]/.test(newPassword)) {
        res.status(400).send("Mật khẩu mới phải chứa ít nhất một chữ cái");
        return;
    }
    if (!/[0-9]/.test(newPassword)) {
        res.status(400).send("Mật khẩu mới phải chứa ít nhất một số");
        return;
    }

    // Hash new password
    let salt = bcrypt.genSaltSync(10);
    let hashedPassword = bcrypt.hashSync(newPassword, salt);

    // Update password
    let userModel = require('../schemas/users');
    await userModel.findByIdAndUpdate(user._id, { password: hashedPassword });

    res.send({ message: "Đổi mật khẩu thành công" });
})


module.exports = router;
