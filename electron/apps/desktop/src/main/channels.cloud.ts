import { cloudBootstrap, cloudLogin, cloudLogout } from "@rpd/shared";
import { CloudService } from "./services/cloud/cloud-service.js";

export function buildCloudHandlers(cloud: CloudService): {
  "cloud/login": (request: ReturnType<typeof cloudLogin["request"]["parse"]>) => ReturnType<typeof cloudLogin["response"]["parse"]> | Promise<ReturnType<typeof cloudLogin["response"]["parse"]>>;
  "cloud/bootstrap": () => ReturnType<typeof cloudBootstrap["response"]["parse"]> | Promise<ReturnType<typeof cloudBootstrap["response"]["parse"]>>;
  "cloud/logout": () => ReturnType<typeof cloudLogout["response"]["parse"]>;
} {
  return {
    "cloud/login": (request) => cloud.login(request.email, request.password),
    "cloud/bootstrap": () => cloud.bootstrap(),
    "cloud/logout": () => {
      cloud.logout();
      return { signedIn: false };
    },
  };
}
