{
  thinkingMode: uiThinkingMode ?? "standard",
  apiThinkingMode,
  multimodalImage:
    typeof multimodalImage === "string" && multimodalImage.length > 0 
      ? multimodalImage 
      : undefined,
}