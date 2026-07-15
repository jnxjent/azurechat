import type {
  HttpMethod,
  HttpRouteHandler,
  IHttpServerAdapter,
  IHttpServerResponse,
} from "@microsoft/teams.apps";

/**
 * Bridges the Teams SDK's framework-neutral HTTP handler to a Next.js route.
 * Next.js owns the HTTP server lifecycle, so start/stop are intentionally absent.
 */
export class NextHttpAdapter implements IHttpServerAdapter {
  private readonly routes = new Map<string, HttpRouteHandler>();

  registerRoute(
    method: HttpMethod,
    path: string,
    handler: HttpRouteHandler
  ): void {
    this.routes.set(this.routeKey(method, path), handler);
  }

  async handle(
    method: HttpMethod,
    path: string,
    request: Request
  ): Promise<IHttpServerResponse> {
    const handler = this.routes.get(this.routeKey(method, path));
    if (!handler) {
      return { status: 404, body: { error: "teams_route_not_found" } };
    }

    let body: unknown;
    try {
      body = await request.json();
    } catch {
      return { status: 400, body: { error: "invalid_json" } };
    }

    return handler({
      body,
      headers: Object.fromEntries(request.headers.entries()),
    });
  }

  private routeKey(method: HttpMethod, path: string): string {
    return `${method} ${path}`;
  }
}
