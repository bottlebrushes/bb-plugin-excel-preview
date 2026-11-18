// `import source from "./x.py" with { type: "text" }`: the plugin build inlines the file as a string.
declare module "*.py" {
  const source: string;
  export default source;
}
