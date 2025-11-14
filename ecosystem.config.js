module.exports = {
  apps: [
    {
      name: "stock-tc-app",
      script: "index.js",
      watch: true, // or specify an array of paths: ["server", "client"]
      ignore_watch: ["node_modules", "logs"],
    },
  ],
};
