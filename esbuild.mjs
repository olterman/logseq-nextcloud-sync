import esbuild from "esbuild";

const watch = process.argv.includes("--watch");

const context = await esbuild.context({
  entryPoints: ["src/main.ts"],
  bundle: true,
  format: "iife",
  platform: "browser",
  target: ["es2020"],
  outfile: "dist/index.js",
  sourcemap: true
});

if (watch) {
  await context.watch();
  console.log("Watching for changes...");
} else {
  await context.rebuild();
  await context.dispose();
  console.log("Built dist/index.js");
}
