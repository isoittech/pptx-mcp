const disabledMessage = "image-size is disabled in PowerPoint MCP; renderer inputs must use verified dimensions.";

export function imageSize() {
  throw new TypeError(disabledMessage);
}

export const disableFS = () => {};
export const disableTypes = () => {};
export const setConcurrency = () => {};
export const types = Object.freeze([]);
export default imageSize;
