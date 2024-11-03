import { Sequelize } from "sequelize";
import { envs } from "./envs";

const sequelize = new Sequelize({
  dialect: "sqlite",
  storage: envs.DB_STORAGE_PATH,
  logging: console.log,
});

export default sequelize;
