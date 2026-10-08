# Excel Merger

The web app runs from a small Express static server. Spreadsheet processing happens in the browser.

## Run on a local network

Run `npm start` (or `npm run dev` for automatic restarts). The server listens on all IPv4 interfaces on port 8080 by default and prints local and network URLs. Open the `Network:` URL for the host's active Wi-Fi or Ethernet address from another device on the same LAN; virtual adapters may also appear in the list. The machine running the server must stay on.

Set `PORT` to use another port, or `HOST` to bind to one interface. For example, setting `HOST` to `127.0.0.1` limits access to the same computer. No machine-specific IP address is stored in the app.

If another device cannot connect, check that both devices are on the same reachable network and allow inbound TCP traffic to the selected port in the host computer's firewall. Some guest Wi-Fi networks block device-to-device traffic. A successful request from the host computer's LAN IP does not establish that another device can reach it.

Anyone who can reach the server can open the app, so use it only on a trusted network. The desktop app (`npm run desktop`) opens local files directly and does not start this web server.
