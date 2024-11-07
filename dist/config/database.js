"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const sequelize_1 = require("sequelize");
const envs_1 = require("./envs");
const sequelize = new sequelize_1.Sequelize({
    dialect: "sqlite",
    storage: envs_1.envs.DB_STORAGE_PATH,
    logging: console.log,
});
exports.default = sequelize;
