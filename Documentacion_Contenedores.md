# Documentacion: Contenedores Docker en AutoMarket Peru

## 1. Que es un Contenedor

Un contenedor es una unidad estandarizada de software que agrupa el codigo y todas sus dependencias para que la aplicacion se ejecute de manera rapida y confiable en cualquier entorno informatico.

A diferencia de una maquina virtual, un contenedor comparte el kernel del sistema operativo host, lo que lo hace mas ligero y rapido de iniciar.

## 2. Beneficios de los Contenedores

- **Portabilidad**: La aplicacion funciona igual en cualquier maquina con Docker instalado
- **Consistencia**: Elimina el problema "en mi maquina funciona"
- **Aislamiento**: Cada servicio corre en su propio contenedor sin afectar a otros
- **Escalabilidad**: Facilita duplicar o escalar servicios segun la demanda
- **CI/CD**: Compatible con GitHub Actions, Kubernetes y cloud providers (AWS, Azure, GCP)

## 3. Arquitectura de Contenedores en AutoMarket Peru

El proyecto utiliza **Docker Compose** para orquestar dos contenedores:

```
┌─────────────────────────────────────────┐
│            docker-compose.yml           │
├──────────────────┬──────────────────────┤
│   FRONTEND       │      BACKEND         │
│   (Nginx)        │   (Python/Flask)     │
│   Puerto 80      │   Puerto 5000        │
├──────────────────┼──────────────────────┤
│   index.html     │      app.py          │
│   styles.css     │   requirements.txt   │
│   incidentes.html│      wsgi.py         │
│   navbar/        │   incidentes.db      │
│   main/          │                      │
└──────────────────┴──────────────────────┘
```

### 3.1 Contenedor Frontend (Nginx)

**Archivo**: `Dockerfile` (raiz del proyecto)

- **Imagen base**: `nginx:alpine` (ligera y optimizada)
- **Funcion**: Sirve los archivos estaticos (HTML, CSS, JS)
- **Puerto**: 80 (accesible desde http://localhost)
- **Configuracion**: `nginx.conf` actua como proxy reverso

### 3.2 Contenedor Backend (Python/Flask)

**Archivo**: `backend/Dockerfile`

- **Imagen base**: `python:3.10-slim`
- **Funcion**: Ejecuta la API REST con Flask y Gunicorn
- **Puerto**: 5000
- **Dependencias**: flask, flask-cors, gunicorn

### 3.3 Orquestacion (Docker Compose)

**Archivo**: `docker-compose.yml`

- Define ambos servicios (frontend y backend)
- Configura puertos y dependencias
- Crea un volumen persistente para datos (`db-data`)

### 3.4 Proxy Reverso (Nginx)

**Archivo**: `nginx.conf`

- Escucha en puerto 80
- Redirige solicitudes `/api/*` al backend en puerto 5000
- Sirve archivos estaticos del frontend

## 4. Cambios Realizados

### 4.1 Correccion del docker-compose.yml

Se elimino la linea `version: "3.8"` porque:

- Es un atributo obsoleto en versiones recientes de Docker Compose
- Generaba un warning al ejecutar `docker-compose up`
- Docker Compose v2+ no requiere esta especificacion

**Antes:**
```yaml
version: "3.8"

services:
  frontend:
    ...
```

**Despues:**
```yaml
services:
  frontend:
    ...
```

## 5. Como Ejecutar el Proyecto

### Prerrequisitos

1. Docker Desktop instalado
2. WSL2 habilitado con distribucion Ubuntu

### Pasos

1. Abrir PowerShell en la carpeta del proyecto
2. Ejecutar:

```powershell
docker-compose up --build
```

3. Abrir el navegador en:
- Frontend: http://localhost
- Backend API: http://localhost:5000

### Comandos Utiles

| Comando | Descripcion |
|---------|-------------|
| `docker-compose up --build` | Construir e iniciar contenedores |
| `docker-compose down` | Detener contenedores |
| `docker-compose ps` | Ver contenedores en ejecucion |
| `docker-compose logs` | Ver logs de los contenedores |
| `docker-compose restart` | Reiniciar contenedores |

## 6. Estructura de Archivos Docker

```
Proyecto_Herramientas/
├── Dockerfile              # Contenedor Frontend
├── docker-compose.yml      # Orquestador
├── nginx.conf              # Configuracion Nginx
├── .dockerignore           # Archivos ignorados
├── backend/
│   ├── Dockerfile          # Contenedor Backend
│   ├── requirements.txt    # Dependencias Python
│   ├── app.py              # Aplicacion Flask
│   └── wsgi.py             # Punto de entrada WSGI
└── ...
```

## 7. Flujo de Datos

```
Usuario (Navegador)
       │
       ▼
   Puerto 80 (Nginx)
       │
       ├──► Archivos Estaticos (HTML/CSS/JS)
       │
       └──► /api/* ──► Puerto 5000 (Flask Backend)
                              │
                              ▼
                      Base de Datos (incidentes.db)
```

---

**Proyecto**: AutoMarket Peru
**Materia**: Herramientas de Desarrollo
**Tema**: Contenedores Docker
**Fecha**: Julio 2026
