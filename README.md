# ICBF — Generador de Informes Nutricionales

Herramienta web para generar informes de seguimiento nutricional a partir de los archivos de tomas exportados del sistema ICBF. Procesa los datos en memoria, genera archivos Excel por contrato y los entrega para descarga individual directamente en el navegador.

---

## Qué hace

A partir de los archivos Excel del sistema ICBF, la app genera por cada contrato un archivo `.xlsx` con tres hojas:

| Hoja | Contenido |
|---|---|
| **Usuarios por Unidad** | Conteo de usuarios únicos vinculados por unidad de servicio |
| **Tomas Faltantes** | Beneficiarios que no han cumplido el intervalo de toma establecido |
| **Alerta Desnutricion / Alerta Nutricional** | Último estado nutricional de cada beneficiario, categorizado y ordenado |

### Estado nutricional — categorías reportadas

**Archivo General** (`ESTADO PESO TALLA`):
- Obesidad
- Sobrepeso
- Riesgo de Sobrepeso
- Peso Adecuado para la Talla
- Riesgo de Desnutrición Aguda
- Desnutrición Aguda Moderada
- Desnutrición Aguda Severa

**Archivo Gestante/Lactante** (`EST.NUTR. GESTANTE`):
- Bajo Peso para la Edad Gestacional
- IMC Adecuado para la Edad Gestacional
- Sobrepeso para la Edad Gestacional
- Obesidad para la Edad Gestacional

---

## Modos de análisis

### Modo Básico
Requiere al menos uno de los dos archivos de nutrición. Los archivos son independientes entre sí.

- Detecta tomas faltantes comparando el historial de tomas registrado contra el intervalo configurado.
- Se puede procesar solo General, solo Gestante/Lactante, o ambos.

### Modo Completo
Requiere los tres archivos. Cruza el historial de tomas con el padrón de beneficiarios activos.

- Detecta beneficiarios sin ninguna toma registrada.
- Calcula tomas esperadas vs. realizadas según la fecha de vinculación.
- Identifica quienes tienen tomas insuficientes respecto al tiempo vinculado.

---

## Archivos de entrada

| Archivo | Descripción | Modo |
|---|---|---|
| `ICBFCUEGeneralPorToma.xlsx` | Tomas nutricionales — población infantil | Básico / Completo |
| `GestanteLactantePorToma.xlsx` | Tomas nutricionales — gestantes y lactantes | Básico / Completo |
| `ICBFCUEBeneficiariosPIActivosRegionalizado.xlsx` | Padrón de beneficiarios activos | Solo Completo |

---

## Estructura del proyecto

```
icbf-app/
├── backend/
│   ├── main.py              # API FastAPI + lógica de procesamiento
│   └── requirements.txt
├── frontend/
│   └── index.html           # Interfaz web (HTML/CSS/JS puro)
├── Dockerfile
├── docker-compose.yml
└── README.md
```

---

## Despliegue en Debian con Docker

### 1. Instalar Docker
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER
# Cerrar y abrir sesión para que aplique
```

### 2. Clonar el repositorio
```bash
git clone https://github.com/migzam10/icfb-informes.git
cd icfb-informes
```

### 3. Construir y levantar
```bash
docker compose up -d --build
```

### 4. Verificar
```bash
docker compose ps
curl http://localhost:8001/health
```

La app queda disponible en: **http://localhost:8001**

---

## Exposición con Cloudflare Tunnel

### Instalar cloudflared
```bash
curl -L https://pkg.cloudflare.com/cloudflare-main.gpg | \
  sudo tee /usr/share/keyrings/cloudflare-main.gpg > /dev/null

echo 'deb [signed-by=/usr/share/keyrings/cloudflare-main.gpg] \
  https://pkg.cloudflare.com/cloudflared bookworm main' | \
  sudo tee /etc/apt/sources.list.d/cloudflared.list

sudo apt update && sudo apt install cloudflared
```

### Crear y configurar el túnel
```bash
cloudflared tunnel login
cloudflared tunnel create icbf-informes
```

`~/.cloudflared/config.yml`:
```yaml
tunnel: icbf-informes
credentials-file: /root/.cloudflared/<UUID>.json

ingress:
  - hostname: informes-icbf.tudominio.com
    service: http://localhost:8001
  - service: http_status:404
```

```bash
cloudflared tunnel route dns icbf-informes informes-icbf.tudominio.com
sudo cloudflared service install
sudo systemctl enable --now cloudflared
```

---

## Comandos útiles

```bash
# Ver logs en tiempo real
docker compose logs -f

# Reiniciar el contenedor
docker compose restart

# Detener
docker compose down

# Reconstruir tras cambios en el código
docker compose up -d --build
```

---

## Notas técnicas

- Todo el procesamiento ocurre en memoria — no se escribe nada en disco ni se necesita base de datos.
- Los archivos generados se entregan directamente al navegador como descarga individual por contrato.
- El tamaño máximo por archivo de entrada es 30 MB.
- El contenedor usa la zona horaria `America/Bogota`.
- Solo procesa beneficiarios con estado `VINCULADO` en el archivo de nutrición.
