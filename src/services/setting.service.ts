import fs from "fs";
import path from "path";
import { Settings } from "../domain/interface/setting.interface";

const settingsFilePath = path.join(
  __dirname,
  "..",
  "../src/data/data/settings.json"
);

export class SettingService {
  constructor() {}

  public async getSettings(): Promise<any> {
    return new Promise((resolve, reject) => {
      fs.readFile(settingsFilePath, "utf8", (err, data) => {
        if (err) {
          console.error("Error al leer el archivo de configuraciones:", err);
          return reject("Error al leer las configuraciones.");
        }
        try {
          const settings = JSON.parse(data);
          resolve(settings);
        } catch (parseError) {
          console.error(
            "Error al parsear el archivo de configuraciones:",
            parseError
          );
          reject("Error al parsear las configuraciones.");
        }
      });
    });
  }

  public async updateSettings(newSettings: Settings): Promise<void> {    
    return new Promise((resolve, reject) => {
      fs.writeFile(
        settingsFilePath,
        JSON.stringify(newSettings, null, 2),
        (err) => {
          if (err) {
            console.error(
              "Error al escribir el archivo de configuraciones:",
              err
            );
            return reject("Error al guardar las configuraciones.");
          }
          resolve();
        }
      );
    });
  }
}
