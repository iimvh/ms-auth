import { driveSignInForm, loadSecrets } from "./adapters/msal/auth.js";
import { providePersonalCaptureDeps } from "./adapters/personal/personal-capture.js";

providePersonalCaptureDeps({ loadSecrets, driveSignInForm });
