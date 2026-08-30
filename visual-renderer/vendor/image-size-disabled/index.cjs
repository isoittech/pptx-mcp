"use strict";

const disabledMessage = "image-size is disabled in PowerPoint MCP; renderer inputs must use verified dimensions.";

function imageSize() {
  throw new TypeError(disabledMessage);
}

imageSize.imageSize = imageSize;
imageSize.default = imageSize;
imageSize.disableFS = () => {};
imageSize.disableTypes = () => {};
imageSize.setConcurrency = () => {};
imageSize.types = Object.freeze([]);

module.exports = imageSize;
