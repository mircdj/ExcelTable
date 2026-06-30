/**
 * Typed event emitter with support for cancellable events.
 * Every grid mutation flows through here, so external code can observe
 * or veto behaviour without subclassing.
 */
import type { EventMap, EventType, GridEventBase } from './types';

type AnyHandler = (e: GridEventBase) => void;

export class Emitter {
  private handlers = new Map<string, Set<AnyHandler>>();

  on<K extends EventType>(type: K, handler: (e: EventMap[K]) => void): () => void {
    let set = this.handlers.get(type);
    if (!set) this.handlers.set(type, (set = new Set()));
    set.add(handler as AnyHandler);
    return () => this.off(type, handler);
  }

  once<K extends EventType>(type: K, handler: (e: EventMap[K]) => void): () => void {
    const off = this.on(type, (e) => {
      off();
      handler(e);
    });
    return off;
  }

  off<K extends EventType>(type: K, handler: (e: EventMap[K]) => void): void {
    this.handlers.get(type)?.delete(handler as AnyHandler);
  }

  /**
   * Emit an event. Returns the full event object so callers can check
   * `defaultPrevented` on cancellable events.
   */
  emit<K extends EventType>(type: K, payload: Omit<EventMap[K], keyof GridEventBase>): EventMap[K] {
    let prevented = false;
    const event = {
      ...payload,
      type,
      timestamp: Date.now(),
      preventDefault: () => {
        prevented = true;
      },
      get defaultPrevented() {
        return prevented;
      },
    } as unknown as EventMap[K];

    const set = this.handlers.get(type);
    if (set) for (const h of [...set]) h(event as GridEventBase);
    return event;
  }

  clear(): void {
    this.handlers.clear();
  }
}
