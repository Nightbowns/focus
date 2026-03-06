# Docker + Cloudflare Tunnel

## Levantar servicios

```powershell
docker compose up -d --build
```

## Ver URL publica del tunel

```powershell
docker compose logs cloudflared
```

Busca la linea que contiene `trycloudflare.com`.

## Verificar app local

```powershell
curl http://localhost:4173
```

## Detener servicios

```powershell
docker compose down
```
