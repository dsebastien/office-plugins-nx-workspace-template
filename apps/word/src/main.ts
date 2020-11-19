import {enableProdMode, NgModuleRef} from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

// Reference: https://www.npmjs.com/package/@microsoft/office-js
if (Office !== undefined) {
  Office.initialize = () => {
    Office.onReady()
      .then((_info: { host: Office.HostType; platform: Office.PlatformType }) => {
        console.log("Office is ready");
        return bootStrapAngular();
      })
      .catch(err => {
        console.warn('Angular not bootstrapped. Error: \n');
        console.log(err);
      });
  };

  Office.initialize(0); // Start
} else {
  console.log('Bootstrapping Angular, without Office JS');
  bootStrapAngular();
}

function bootStrapAngular(): Promise<void | NgModuleRef<AppModule>> {
  return platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(error => {
      console.log('Error bootstrapping Office Angular app: ', error);
    });
}
