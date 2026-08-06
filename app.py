<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cargando...</title>
  <style>
    :root {
      --bg-color: #0f172a;
      --card-bg: #1e293b;
      --accent: #6366f1;
      --accent-glow: rgba(99, 102, 241, 0.35);
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    body {
      background-color: var(--bg-color);
      color: var(--text-main);
      height: 100vh;
      width: 100vw;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
    }

    .loader-card {
      background: var(--card-bg);
      padding: 2.5rem 3rem;
      border-radius: 1.25rem;
      border: 1px solid rgba(255, 255, 255, 0.08);
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4);
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 1.75rem;
      max-width: 380px;
      width: 90%;
      text-align: center;
    }

    /* Indicador central animado */
    .spinner-box {
      position: relative;
      width: 70px;
      height: 70px;
      display: flex;
      align-items: center;
      justify-content: center;
    }

    .circle-outer {
      position: absolute;
      width: 100%;
      height: 100%;
      border-radius: 50%;
      border: 3px solid transparent;
      border-top-color: var(--accent);
      border-right-color: var(--accent);
      animation: spin 1.2s cubic-bezier(0.5, 0, 0.5, 1) infinite;
      filter: drop-shadow(0 0 8px var(--accent-glow));
    }

    .circle-inner {
      position: absolute;
      width: 60%;
      height: 60%;
      border-radius: 50%;
      border: 2px solid rgba(255, 255, 255, 0.1);
    }

    .pulse-dot {
      width: 12px;
      height: 12px;
      background-color: var(--accent);
      border-radius: 50%;
      animation: pulse 1.5s ease-in-out infinite;
    }

    /* Mensajes y tipografía */
    .content-box {
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }

    .title {
      font-size: 1.15rem;
      font-weight: 600;
      letter-spacing: -0.01em;
    }

    .subtitle {
      font-size: 0.875rem;
      color: var(--text-muted);
      line-height: 1.4;
    }

    /* Barra de progreso minimalista */
    .progress-bar-container {
      width: 100%;
      height: 4px;
      background: rgba(255, 255, 255, 0.08);
      border-radius: 2px;
      overflow: hidden;
      position: relative;
    }

    .progress-bar-value {
      width: 45%;
      height: 100%;
      background: var(--accent);
      border-radius: 2px;
      position: absolute;
      animation: indeterminate 2s ease-in-out infinite;
    }

    /* Animaciones */
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    @keyframes pulse {
      0%, 100% { transform: scale(0.8); opacity: 0.5; }
      50% { transform: scale(1.2); opacity: 1; }
    }

    @keyframes indeterminate {
      0% { left: -35%; width: 30%; }
      50% { left: 35%; width: 60%; }
      100% { left: 100%; width: 30%; }
    }
  </style>
</head>
<body>

  <main class="loader-card">
    <div class="spinner-box">
      <div class="circle-outer"></div>
      <div class="circle-inner"></div>
      <div class="pulse-dot"></div>
    </div>

    <div class="content-box">
      <h1 class="title">Cargando aplicación</h1>
      <p class="subtitle">Sincronizando datos del sistema, por favor espera un momento...</p>
    </div>

    <div class="progress-bar-container">
      <div class="progress-bar-value"></div>
    </div>
  </main>

</body>
</html>
