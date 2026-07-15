# Docker 部署说明

## 构建并启动

在项目根目录执行：

```bash
docker compose up -d --build
```

启动后访问：

```text
http://服务器IP:8000
```

前端页面和后端接口都由同一个容器提供：

- 页面：`/`
- 健康检查：`/api/health`
- 上传/预览/转换/下载接口：`/api/...`

## 查看日志

```bash
docker compose logs -f webword
```

## 停止服务

```bash
docker compose down
```

## 更新部署

拉取或拷贝新代码后重新构建：

```bash
docker compose up -d --build
```

## 端口调整

默认映射为：

```yaml
ports:
  - "8000:8000"
```

如果服务器 8000 被占用，可以把左侧端口改成其他端口，例如：

```yaml
ports:
  - "8080:8000"
```

然后访问 `http://服务器IP:8080`。
