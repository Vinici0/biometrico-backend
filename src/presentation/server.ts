import express, { Router } from "express";
import cors from "cors";
import path from "path";
import morgan from 'morgan';

import sequelize from "../config/database";

interface Options {
  port: number;
  routes: Router;
  public_path?: string;
}

export class Server {
  public readonly app = express();
  private serverListener?: any;
  private readonly port: number;
  private readonly publicPath: string;
  private readonly routes: Router;

  constructor(options: Options) {
    const { port, public_path = "public" } = options;
    this.port = port;
    this.publicPath = public_path;
    this.routes = options.routes;
    this.configure();
  }

  private configure() {
    //* CORS
    this.app.use(cors());

    //* Middlewares
    this.app.use(express.json());
    this.app.use(express.urlencoded({ extended: true }));

    //* Public Folder
    this.app.use(express.static(this.publicPath));

    this.app.get(/^\/(?!api).*/, (req, res) => {
      const indexPath = path.join(
        __dirname + `../../../${this.publicPath}/index.html`
      );
      res.sendFile(indexPath);
    });

    //* Logger
    // this.app.use(morgan('dev'));
    this.app.use(morgan('combined'));

    //* Database
    this.connectDatabase();

    //* Routes
    this.setRoutes(this.routes);
  }

  public setRoutes(router: Router) {
    this.app.use(router);
  }

  public async connectDatabase() {
    try {
      // await sequelize.authenticate();
      sequelize.sync({ force: false }).then(() => {});
      console.log("Database connected");
    } catch (error) {
      console.log("Error: ", error);
    }
  }

  async start() {
    this.serverListener = this.app.listen(this.port, () => {
      console.log(`Server running on port ${this.port}`);
    });
  }

  public close() {
    this.serverListener?.close();
  }
}
