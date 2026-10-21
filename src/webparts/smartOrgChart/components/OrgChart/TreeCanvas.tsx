import * as React from 'react';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './OrgChart.module.scss';
import { IChartCamera } from './chartPersistence';

export type ICanvasCamera = Omit<IChartCamera, 'focus'>;

interface ITreeCanvasProps {
  /** Which way the tree grows: the root sits at the top (vertical) or at the left (horizontal). */
  layout: 'vertical' | 'horizontal';
  defaultZoom: number;
  className: string;
  busy: boolean;
  /** Where to put the camera when the canvas first mounts; fits the tree when absent or not found. */
  initialCamera?: ICanvasCamera;
  /** Reported shortly after the camera stops moving, and when the canvas unmounts. */
  onCamera?: (camera: ICanvasCamera) => void;
  onZoom: (zoom: number) => void;
  onViewport: (el: HTMLDivElement | null) => void;
}

// Camera limits and gesture tuning
export const MIN_ZOOM = 0.01;
export const MAX_ZOOM = 1.5;
const FIT_PADDING_PX = 48;          // breathing room around the chart in "fit"
const OPEN_MIN_ZOOM = 0.6;          // a tree too big to fit at this zoom opens at it, anchored at the root
const OPEN_MARGIN_PX = 24;          // gap between the root card and the viewport edge in that view
const DRAG_THRESHOLD_PX = 5;        // movement before a press becomes a pan (keeps taps working)
const CLICK_SUPPRESS_MS = 350;      // swallow the click that ends a drag
const WHEEL_SENSITIVITY = 0.002;    // zoom factor exponent per wheel delta unit
const WHEEL_MAX_DELTA = 100;        // clamp so one fast flick can't jump the zoom
const WHEEL_LINE_PX = 16;           // deltaMode 1 (lines) → pixels
const CAMERA_REPORT_MS = 400;       // quiet time before the camera position is reported

export const clampZoom = (zoom: number): number => Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, zoom));

/** A DOM tree on a map-like camera. Camera updates never re-render its cards. */
export class TreeCanvas extends React.Component<ITreeCanvasProps, { active: boolean }> {
  public state = { active: false };
  private viewport = React.createRef<HTMLDivElement>();
  private content = React.createRef<HTMLDivElement>();
  private x = 0;
  private y = 0;
  private zoom = 1;
  private automatic = true;
  private observer: ResizeObserver | undefined;
  private frame = 0;
  private cameraTimer = 0;
  private pointers = new Map<number, { x: number; y: number }>();
  private gesture: { x: number; y: number; distance: number } | undefined;
  private moved = false;
  private suppressClickUntil = 0;
  private touchQuery = window.matchMedia('(pointer: coarse)');

  public componentDidMount(): void {
    const el = this.viewport.current;
    const content = this.content.current;
    if (!el || !content) return;
    this.props.onViewport(el);
    el.addEventListener('wheel', this.wheel, { passive: false });
    this.observer = new ResizeObserver(() => {
      if (this.automatic) this.openView();
    });
    this.observer.observe(el);
    this.observer.observe(content);
    const initial = this.props.initialCamera;
    if (!(initial && !initial.auto && this.restore(initial))) this.reset();
  }

  public componentDidUpdate(prev: ITreeCanvasProps): void {
    if (prev.defaultZoom !== this.props.defaultZoom) this.reset();
  }

  public componentWillUnmount(): void {
    this.viewport.current?.removeEventListener('wheel', this.wheel);
    this.observer?.disconnect();
    cancelAnimationFrame(this.frame);
    if (this.cameraTimer) {
      window.clearTimeout(this.cameraTimer);
      this.cameraTimer = 0;
      this.reportCamera();   // don't lose a move made just before leaving the view
    }
    this.props.onViewport(null);
  }

  // Center of a card in the content's own (unscaled) coordinates. Offsets ignore
  // the camera transform, so this works at any pan/zoom.
  private cardCenter(card: HTMLElement): { x: number; y: number } | undefined {
    const content = this.content.current;
    if (!content) return undefined;
    let x = card.offsetWidth / 2;
    let y = card.offsetHeight / 2;
    let el: HTMLElement | null = card;
    while (el && el !== content) {
      x += el.offsetLeft;
      y += el.offsetTop;
      el = el.offsetParent as HTMLElement | null;
    }
    return el === content ? { x, y } : undefined;
  }

  private currentCamera(): ICanvasCamera | undefined {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c) return undefined;
    if (this.automatic) return { auto: true, zoom: this.zoom, id: '', dx: 0, dy: 0 };
    const cx = (v.clientWidth / 2 - this.x) / this.zoom;
    const cy = (v.clientHeight / 2 - this.y) / this.zoom;
    let best: { id: string; x: number; y: number; d: number } | undefined;
    const cards = c.querySelectorAll('[data-soc-id]');
    for (let i = 0; i < cards.length; i++) {
      const center = this.cardCenter(cards[i] as HTMLElement);
      if (!center) continue;
      const d = (center.x - cx) * (center.x - cx) + (center.y - cy) * (center.y - cy);
      if (!best || d < best.d) best = { id: cards[i].getAttribute('data-soc-id') || '', x: center.x, y: center.y, d };
    }
    if (!best || !best.id) return undefined;
    return {
      auto: false, zoom: this.zoom, id: best.id,
      dx: best.x * this.zoom + this.x - v.clientWidth / 2,
      dy: best.y * this.zoom + this.y - v.clientHeight / 2,
    };
  }

  private reportCamera(): void {
    const camera = this.props.onCamera && this.currentCamera();
    if (camera && this.props.onCamera) this.props.onCamera(camera);
  }

  /** Puts the camera back where a saved view left it. False (nothing changed) if its anchor card isn't in the tree. */
  public restore = (camera: ICanvasCamera): boolean => {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c || camera.auto || !camera.id) return false;
    const cards = c.querySelectorAll('[data-soc-id]');
    let anchor: HTMLElement | undefined;
    for (let i = 0; i < cards.length; i++) {
      if (cards[i].getAttribute('data-soc-id') === camera.id) { anchor = cards[i] as HTMLElement; break; }
    }
    const center = anchor && this.cardCenter(anchor);
    if (!center) return false;
    this.zoom = clampZoom(camera.zoom);
    this.x = v.clientWidth / 2 + camera.dx - center.x * this.zoom;
    this.y = v.clientHeight / 2 + camera.dy - center.y * this.zoom;
    this.automatic = false;
    this.paint();
    this.props.onZoom(this.zoom);
    return true;
  };

  private paint(): void {
    if (this.cameraTimer) window.clearTimeout(this.cameraTimer);
    if (this.props.onCamera) {
      this.cameraTimer = window.setTimeout(() => { this.cameraTimer = 0; this.reportCamera(); }, CAMERA_REPORT_MS);
    }
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      if (this.content.current) {
        this.content.current.style.transform = `translate(${this.x}px, ${this.y}px) scale(${this.zoom})`;
      }
    });
  }

  private center(zoom: number): void {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c) return;
    this.zoom = clampZoom(zoom);
    this.x = (v.clientWidth - c.offsetWidth * this.zoom) / 2;
    this.y = (v.clientHeight - c.offsetHeight * this.zoom) / 2;
    this.paint();
    this.props.onZoom(this.zoom);
  }

  public reset = (): void => {
    this.automatic = !(this.props.defaultZoom > 0);
    if (this.automatic) this.openView();
    else this.center(this.props.defaultZoom);
  };

  // The zoom at which the whole tree fits the viewport (never above 100%)
  private fitZoom(): number | undefined {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c || !c.offsetWidth || !c.offsetHeight) return undefined;
    return Math.min(1, Math.max(1, v.clientWidth - FIT_PADDING_PX) / c.offsetWidth,
      Math.max(1, v.clientHeight - FIT_PADDING_PX) / c.offsetHeight);
  }

  // How the chart opens. A tree that fits at a readable zoom is simply fitted. A bigger one
  // opens at that readable zoom, looking at its root, instead of shrinking to an unreadable
  // overview — **Fit entire chart** is still one click away. Either way the view stays
  // automatic — re-applied as the tree loads, grows or the window resizes, so the root stays
  // put — until the user pans, zooms or picks a view themselves.
  private openView(): void {
    const v = this.viewport.current;
    const c = this.content.current;
    const fit = this.fitZoom();
    if (!v || !c || fit === undefined) return;
    const root = c.querySelector('[data-soc-id]') as HTMLElement | null;
    const rootCenter = root && this.cardCenter(root);
    if (fit >= OPEN_MIN_ZOOM || !root || !rootCenter) {
      this.center(fit);
      return;
    }
    this.zoom = OPEN_MIN_ZOOM;
    // The root card sits at the top-left of a left-to-right tree and top-center of a top-down one
    this.x = this.props.layout === 'horizontal'
      ? OPEN_MARGIN_PX - (rootCenter.x - root.offsetWidth / 2) * this.zoom
      : v.clientWidth / 2 - rootCenter.x * this.zoom;
    this.y = OPEN_MARGIN_PX - (rootCenter.y - root.offsetHeight / 2) * this.zoom;
    this.paint();
    this.props.onZoom(this.zoom);
  }

  /** Zooms out until the whole tree is in view, however small that makes it. */
  public fit = (): void => {
    const fit = this.fitZoom();
    if (fit === undefined) return;
    this.automatic = false;   // an explicit choice: expanding a branch or resizing leaves it alone
    this.center(fit);
  };

  public autoFit = (): void => { if (this.automatic) this.openView(); };

  public setZoom = (zoom: number, anchorX?: number, anchorY?: number): void => {
    const v = this.viewport.current;
    if (!v || !Number.isFinite(zoom)) return;
    const next = clampZoom(zoom);
    const ax = anchorX === undefined ? v.clientWidth / 2 : anchorX;
    const ay = anchorY === undefined ? v.clientHeight / 2 : anchorY;
    this.x = ax - (ax - this.x) * next / this.zoom;
    this.y = ay - (ay - this.y) * next / this.zoom;
    this.zoom = next;
    this.automatic = false;
    this.paint();
    this.props.onZoom(next);
  };

  public reveal = (card: HTMLElement): void => {
    // Flush pending transforms before measuring screen coordinates.
    const c = this.content.current;
    const v = this.viewport.current;
    if (!v || !c) return;
    c.style.transform = `translate(${this.x}px, ${this.y}px) scale(${this.zoom})`;
    const a = card.getBoundingClientRect();
    const b = v.getBoundingClientRect();
    this.x += b.left + b.width / 2 - (a.left + a.width / 2);
    this.y += b.top + b.height / 2 - (a.top + a.height / 2);
    this.automatic = false;
    this.paint();
  };

  private wheel = (e: WheelEvent): void => {
    if (this.touchQuery.matches && !this.state.active) return;
    if ((e.target as HTMLElement).closest('button, input, a')) return;
    e.preventDefault();
    const viewport = this.viewport.current;
    if (!viewport) return;
    const r = viewport.getBoundingClientRect();
    const delta = e.deltaY * (e.deltaMode === 1 ? WHEEL_LINE_PX : e.deltaMode === 2 ? r.height : 1);
    this.setZoom(this.zoom * Math.exp(-Math.max(-WHEEL_MAX_DELTA, Math.min(WHEEL_MAX_DELTA, delta)) * WHEEL_SENSITIVITY), e.clientX - r.left, e.clientY - r.top);
  };

  private measureGesture(): { x: number; y: number; distance: number } {
    const points: Array<{ x: number; y: number }> = [];
    this.pointers.forEach(p => points.push(p));
    const a = points[0];
    const b = points[1] || a;
    return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2,
      distance: Math.sqrt((a.x - b.x) ** 2 + (a.y - b.y) ** 2) };
  }

  private pointerDown = (e: React.PointerEvent<HTMLDivElement>): void => {
    if ((e.pointerType !== 'mouse' || this.touchQuery.matches) && !this.state.active) return;
    if (e.button !== 0 || (e.target as HTMLElement).closest('button, a, input')) return;
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    this.gesture = this.measureGesture();
    if (this.pointers.size === 1) this.moved = false;
    // Capture only after a drag threshold, so taps still reach the card.
  };

  private pointerMove = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!this.pointers.has(e.pointerId) || !this.gesture) return;
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    const next = this.measureGesture();
    const dx = next.x - this.gesture.x;
    const dy = next.y - this.gesture.y;
    if (!this.moved && this.pointers.size < 2 && Math.sqrt(dx * dx + dy * dy) < DRAG_THRESHOLD_PX) return;
    e.currentTarget.setPointerCapture(e.pointerId);
    this.moved = true;
    this.automatic = false;
    if (this.gesture.distance > 0 && next.distance > 0) {
      const viewport = this.viewport.current;
      if (!viewport) return;
      const r = viewport.getBoundingClientRect();
      this.setZoom(this.zoom * next.distance / this.gesture.distance, this.gesture.x - r.left, this.gesture.y - r.top);
    }
    this.x += dx;
    this.y += dy;
    this.gesture = next;
    this.paint();
    e.currentTarget.style.cursor = 'grabbing';
  };

  private pointerEnd = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!this.pointers.delete(e.pointerId)) return;
    if (this.moved) this.suppressClickUntil = Date.now() + CLICK_SUPPRESS_MS;
    if (e.currentTarget.hasPointerCapture(e.pointerId)) e.currentTarget.releasePointerCapture(e.pointerId);
    this.gesture = this.pointers.size ? this.measureGesture() : undefined;
    if (!this.pointers.size) e.currentTarget.style.cursor = '';
  };

  private deactivate = (): void => {
    const v = this.viewport.current;
    this.pointers.forEach((_p, id) => { if (v?.hasPointerCapture(id)) v.releasePointerCapture(id); });
    this.pointers.clear();
    this.gesture = undefined;
    if (v) v.style.cursor = '';
    this.setState({ active: false });
  };

  private focusVisible = (e: React.FocusEvent<HTMLDivElement>): void => {
    const card = (e.target as HTMLElement).closest('[data-soc-id]') as HTMLElement | null;
    const v = this.viewport.current;
    if (!card || !v) return;
    const a = card.getBoundingClientRect();
    const b = v.getBoundingClientRect();
    if ((e.target as HTMLElement).matches(':focus-visible') &&
      (a.left < b.left || a.right > b.right || a.top < b.top || a.bottom > b.bottom)) this.reveal(card);
    // Native focus scrolling must not create a second, invisible camera.
    requestAnimationFrame(() => { if (this.viewport.current) {
      this.viewport.current.scrollLeft = 0;
      this.viewport.current.scrollTop = 0;
    } });
  };

  public render(): React.ReactElement {
    return <div ref={this.viewport}
      className={`${this.props.className} ${styles.treeCanvas} ${this.state.active ? styles.canvasActive : ''}`}
      aria-busy={this.props.busy}
      onPointerDown={this.pointerDown} onPointerMove={this.pointerMove}
      onPointerUp={this.pointerEnd} onPointerCancel={this.pointerEnd}
      // A touch that starts on a card has implicit capture on that card; setting our own capture
      // releases it, and that event bubbles up here. Only the loss of OUR capture ends the drag.
      onLostPointerCapture={e => { if (e.target === e.currentTarget) this.pointerEnd(e); }}
      onPointerLeave={e => { if (!e.currentTarget.hasPointerCapture(e.pointerId)) this.pointerEnd(e); }}
      onKeyDown={e => { if (e.key === 'Escape') this.deactivate(); }}
      onFocusCapture={this.focusVisible}
      onDragStart={e => e.preventDefault()}
      onClickCapture={e => {
        if (!(e.target as HTMLElement).closest(`.${styles.canvasTouchControls}`) && Date.now() < this.suppressClickUntil) {
          e.preventDefault(); e.stopPropagation();
        }
      }}>
      <div ref={this.content} className={styles.canvasContent}>{this.props.children}</div>
      <div className={styles.canvasTouchControls}>
        <button type="button" aria-pressed={this.state.active}
          onClick={this.state.active ? this.deactivate : () => this.setState({ active: true })}>
          {this.state.active ? strings.Chart_CanvasDone : strings.Chart_CanvasActivate}
        </button>
        {!this.state.active && <span>{strings.Chart_CanvasHint}</span>}
      </div>
      {!this.state.active && <div className={styles.canvasTouchShield} aria-hidden="true" />}
    </div>;
  }
}
