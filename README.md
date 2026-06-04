# ICBF — Informes Nutricionales

App web para generar informes de tomas faltantes y alertas nutricionales.
Backend: FastAPI + Python. Frontend: HTML/CSS/JS puro. Contenedor: Docker.

---

## Estructura del proyecto

```
icbf-app/
├── backend/
│   ├── main.py              # API FastAPI + lógica de informes
│   └── requirements.txt
├── frontend/
│   └── index.html           # Interfaz web (HTML puro)
├── Dockerfile
├── docker-compose.yml
└── README.md
```

---

## Despliegue en Debian con Docker

### 1. Instalar Docker (si no está)
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER
# Cerrar y abrir sesión para que aplique
```

### 2. Clonar/copiar el proyecto
```bash
# Copiar la carpeta icbf-app a tu servidor
scp -r icbf-app/ usuario@tu-servidor:/opt/icbf-app
# O con git si lo subes a un repo
```

### 3. Construir y levantar
```bash
cd /opt/icbf-app
docker compose up -d --build
```

### 4. Verificar que está corriendo
```bash
docker compose ps
curl http://localhost:8001/health
```

La app queda disponible en: **http://localhost:8001**

---

## Exposición con Cloudflare Tunnel

### Instalar cloudflared
```bash
# En Debian
curl -L https://pkg.cloudflare.com/cloudflare-main.gpg | \
  sudo tee /usr/share/keyrings/cloudflare-main.gpg > /dev/null

echo 'deb [signed-by=/usr/share/keyrings/cloudflare-main.gpg] \
  https://pkg.cloudflare.com/cloudflared bookworm main' | \
  sudo tee /etc/apt/sources.list.d/cloudflared.list

sudo apt update && sudo apt install cloudflared
```

### Autenticar y crear el túnel
```bash
cloudflared tunnel login
cloudflared tunnel create icbf-informes
```

### Configurar el túnel (~/.cloudflared/config.yml)
```yaml
tunnel: icbf-informes
credentials-file: /root/.cloudflared/<UUID>.json

ingress:
  - hostname: informes-icbf.tudominio.com
    service: http://localhost:8001
  - service: http_status:404
```

### Crear registro DNS
```bash
cloudflared tunnel route dns icbf-informes informes-icbf.tudominio.com
```

### Instalar como servicio systemd
```bash
sudo cloudflared service install
sudo systemctl enable cloudflared
sudo systemctl start cloudflared
```

---

## Comandos útiles

```bash
# Ver logs en tiempo real
docker compose logs -f

# Reiniciar
docker compose restart

# Detener
docker compose down

# Reconstruir tras cambios
docker compose up -d --build
```

---

## Notas importantes

- Los archivos subidos se almacenan en tmpfs (RAM) y se eliminan automáticamente.
  No persisten entre reinicios del contenedor. No se necesita base de datos.
- El contenedor usa la zona horaria America/Bogota.
- Si el servidor se apaga, la app no estará disponible. Para alta disponibilidad
  considera Hugging Face Spaces o un VPS.
