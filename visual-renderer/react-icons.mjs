import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import {
  LuArchiveRestore,
  LuBot,
  LuChartNoAxesCombined,
  LuCircleCheck,
  LuClock3,
  LuCloud,
  LuDatabase,
  LuFileText,
  LuLandmark,
  LuLifeBuoy,
  LuLightbulb,
  LuLockKeyhole,
  LuMessagesSquare,
  LuMonitor,
  LuNetwork,
  LuScale,
  LuSearch,
  LuSettings,
  LuShieldCheck,
  LuTarget,
  LuTriangleAlert,
  LuUsers,
} from "react-icons/lu";

export const approvedReactIcons = Object.freeze({
  insight: LuLightbulb,
  target: LuTarget,
  growth: LuChartNoAxesCombined,
  people: LuUsers,
  shield: LuShieldCheck,
  clock: LuClock3,
  cloud: LuCloud,
  settings: LuSettings,
  data: LuDatabase,
  warning: LuTriangleAlert,
  check: LuCircleCheck,
  idea: LuLightbulb,
  search: LuSearch,
  compliance: LuScale,
  decision: LuTarget,
  lock: LuLockKeyhole,
  network: LuNetwork,
  document: LuFileText,
  communication: LuMessagesSquare,
  recovery: LuLifeBuoy,
  backup: LuArchiveRestore,
  legal: LuLandmark,
  monitor: LuMonitor,
  automation: LuBot,
});

export function renderApprovedReactIcon(iconName, options = {}) {
  const normalized = String(iconName ?? "").trim().toLowerCase();
  const Icon = approvedReactIcons[normalized];
  if (!Icon) {
    throw new Error(`Unsupported react-icons identifier: ${normalized || "(empty)"}`);
  }

  return renderToStaticMarkup(createElement(Icon, {
    "aria-hidden": "true",
    color: options.color ?? "currentColor",
    size: options.size ?? 96,
    strokeWidth: options.strokeWidth ?? 1.8,
    style: {
      display: "block",
      height: "100%",
      width: "100%",
    },
  }));
}
