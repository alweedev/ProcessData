import { afterEach, describe, expect, it, vi } from "vitest";
import { postFormForBlob } from "./api";

class FakeXHR {
  static last: FakeXHR;
  upload = new EventTarget();
  responseType = "";
  status = 200;
  response: Blob | null = null;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  constructor() {
    FakeXHR.last = this;
  }
  open() {}
  send() {}
}

describe("postFormForBlob — progresso do envio", () => {
  afterEach(() => vi.unstubAllGlobals());

  function start(onProgress: (pct: number) => void) {
    vi.stubGlobal("XMLHttpRequest", FakeXHR);
    void postFormForBlob("/api/x", new FormData(), onProgress).catch(() => {});
    return FakeXHR.last;
  }

  it("converte loaded/total dos eventos de progresso em porcentagem", () => {
    const onProgress = vi.fn();
    const xhr = start(onProgress);
    xhr.upload.dispatchEvent(new ProgressEvent("progress", { lengthComputable: true, loaded: 50, total: 200 }));
    expect(onProgress).toHaveBeenLastCalledWith(25);
  });

  it("informa 100% quando o envio termina, mesmo sem nenhum evento de progresso (arquivo pequeno)", () => {
    const onProgress = vi.fn();
    const xhr = start(onProgress);
    xhr.upload.dispatchEvent(new Event("load"));
    expect(onProgress).toHaveBeenLastCalledWith(100);
  });
});
