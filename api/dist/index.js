var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });
var __esm = (fn, res) => function __init() {
  return fn && (res = (0, fn[__getOwnPropNames(fn)[0]])(fn = 0)), res;
};
var __commonJS = (cb, mod) => function __require() {
  return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
};
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __publicField = (obj, key, value) => {
  __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);
  return value;
};

// ../node_modules/unenv/dist/runtime/_internal/utils.mjs
function createNotImplementedError(name) {
  return new Error(`[unenv] ${name} is not implemented yet!`);
}
function notImplemented(name) {
  const fn = /* @__PURE__ */ __name(() => {
    throw createNotImplementedError(name);
  }, "fn");
  return Object.assign(fn, { __unenv__: true });
}
function notImplementedClass(name) {
  return class {
    __unenv__ = true;
    constructor() {
      throw new Error(`[unenv] ${name} is not implemented yet!`);
    }
  };
}
var init_utils = __esm({
  "../node_modules/unenv/dist/runtime/_internal/utils.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    __name(createNotImplementedError, "createNotImplementedError");
    __name(notImplemented, "notImplemented");
    __name(notImplementedClass, "notImplementedClass");
  }
});

// ../node_modules/unenv/dist/runtime/node/internal/perf_hooks/performance.mjs
var _timeOrigin, _performanceNow, nodeTiming, PerformanceEntry, PerformanceMark, PerformanceMeasure, PerformanceResourceTiming, PerformanceObserverEntryList, Performance, PerformanceObserver, performance;
var init_performance = __esm({
  "../node_modules/unenv/dist/runtime/node/internal/perf_hooks/performance.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_utils();
    _timeOrigin = globalThis.performance?.timeOrigin ?? Date.now();
    _performanceNow = globalThis.performance?.now ? globalThis.performance.now.bind(globalThis.performance) : () => Date.now() - _timeOrigin;
    nodeTiming = {
      name: "node",
      entryType: "node",
      startTime: 0,
      duration: 0,
      nodeStart: 0,
      v8Start: 0,
      bootstrapComplete: 0,
      environment: 0,
      loopStart: 0,
      loopExit: 0,
      idleTime: 0,
      uvMetricsInfo: {
        loopCount: 0,
        events: 0,
        eventsWaiting: 0
      },
      detail: void 0,
      toJSON() {
        return this;
      }
    };
    PerformanceEntry = class {
      __unenv__ = true;
      detail;
      entryType = "event";
      name;
      startTime;
      constructor(name, options2) {
        this.name = name;
        this.startTime = options2?.startTime || _performanceNow();
        this.detail = options2?.detail;
      }
      get duration() {
        return _performanceNow() - this.startTime;
      }
      toJSON() {
        return {
          name: this.name,
          entryType: this.entryType,
          startTime: this.startTime,
          duration: this.duration,
          detail: this.detail
        };
      }
    };
    __name(PerformanceEntry, "PerformanceEntry");
    PerformanceMark = /* @__PURE__ */ __name(class PerformanceMark2 extends PerformanceEntry {
      entryType = "mark";
      constructor() {
        super(...arguments);
      }
      get duration() {
        return 0;
      }
    }, "PerformanceMark");
    PerformanceMeasure = class extends PerformanceEntry {
      entryType = "measure";
    };
    __name(PerformanceMeasure, "PerformanceMeasure");
    PerformanceResourceTiming = class extends PerformanceEntry {
      entryType = "resource";
      serverTiming = [];
      connectEnd = 0;
      connectStart = 0;
      decodedBodySize = 0;
      domainLookupEnd = 0;
      domainLookupStart = 0;
      encodedBodySize = 0;
      fetchStart = 0;
      initiatorType = "";
      name = "";
      nextHopProtocol = "";
      redirectEnd = 0;
      redirectStart = 0;
      requestStart = 0;
      responseEnd = 0;
      responseStart = 0;
      secureConnectionStart = 0;
      startTime = 0;
      transferSize = 0;
      workerStart = 0;
      responseStatus = 0;
    };
    __name(PerformanceResourceTiming, "PerformanceResourceTiming");
    PerformanceObserverEntryList = class {
      __unenv__ = true;
      getEntries() {
        return [];
      }
      getEntriesByName(_name, _type) {
        return [];
      }
      getEntriesByType(type) {
        return [];
      }
    };
    __name(PerformanceObserverEntryList, "PerformanceObserverEntryList");
    Performance = class {
      __unenv__ = true;
      timeOrigin = _timeOrigin;
      eventCounts = /* @__PURE__ */ new Map();
      _entries = [];
      _resourceTimingBufferSize = 0;
      navigation = void 0;
      timing = void 0;
      timerify(_fn, _options) {
        throw createNotImplementedError("Performance.timerify");
      }
      get nodeTiming() {
        return nodeTiming;
      }
      eventLoopUtilization() {
        return {};
      }
      markResourceTiming() {
        return new PerformanceResourceTiming("");
      }
      onresourcetimingbufferfull = null;
      now() {
        if (this.timeOrigin === _timeOrigin) {
          return _performanceNow();
        }
        return Date.now() - this.timeOrigin;
      }
      clearMarks(markName) {
        this._entries = markName ? this._entries.filter((e) => e.name !== markName) : this._entries.filter((e) => e.entryType !== "mark");
      }
      clearMeasures(measureName) {
        this._entries = measureName ? this._entries.filter((e) => e.name !== measureName) : this._entries.filter((e) => e.entryType !== "measure");
      }
      clearResourceTimings() {
        this._entries = this._entries.filter((e) => e.entryType !== "resource" || e.entryType !== "navigation");
      }
      getEntries() {
        return this._entries;
      }
      getEntriesByName(name, type) {
        return this._entries.filter((e) => e.name === name && (!type || e.entryType === type));
      }
      getEntriesByType(type) {
        return this._entries.filter((e) => e.entryType === type);
      }
      mark(name, options2) {
        const entry = new PerformanceMark(name, options2);
        this._entries.push(entry);
        return entry;
      }
      measure(measureName, startOrMeasureOptions, endMark) {
        let start;
        let end;
        if (typeof startOrMeasureOptions === "string") {
          start = this.getEntriesByName(startOrMeasureOptions, "mark")[0]?.startTime;
          end = this.getEntriesByName(endMark, "mark")[0]?.startTime;
        } else {
          start = Number.parseFloat(startOrMeasureOptions?.start) || this.now();
          end = Number.parseFloat(startOrMeasureOptions?.end) || this.now();
        }
        const entry = new PerformanceMeasure(measureName, {
          startTime: start,
          detail: {
            start,
            end
          }
        });
        this._entries.push(entry);
        return entry;
      }
      setResourceTimingBufferSize(maxSize) {
        this._resourceTimingBufferSize = maxSize;
      }
      addEventListener(type, listener, options2) {
        throw createNotImplementedError("Performance.addEventListener");
      }
      removeEventListener(type, listener, options2) {
        throw createNotImplementedError("Performance.removeEventListener");
      }
      dispatchEvent(event) {
        throw createNotImplementedError("Performance.dispatchEvent");
      }
      toJSON() {
        return this;
      }
    };
    __name(Performance, "Performance");
    PerformanceObserver = class {
      __unenv__ = true;
      _callback = null;
      constructor(callback) {
        this._callback = callback;
      }
      takeRecords() {
        return [];
      }
      disconnect() {
        throw createNotImplementedError("PerformanceObserver.disconnect");
      }
      observe(options2) {
        throw createNotImplementedError("PerformanceObserver.observe");
      }
      bind(fn) {
        return fn;
      }
      runInAsyncScope(fn, thisArg, ...args) {
        return fn.call(thisArg, ...args);
      }
      asyncId() {
        return 0;
      }
      triggerAsyncId() {
        return 0;
      }
      emitDestroy() {
        return this;
      }
    };
    __name(PerformanceObserver, "PerformanceObserver");
    __publicField(PerformanceObserver, "supportedEntryTypes", []);
    performance = globalThis.performance && "addEventListener" in globalThis.performance ? globalThis.performance : new Performance();
  }
});

// ../node_modules/unenv/dist/runtime/node/perf_hooks.mjs
var init_perf_hooks = __esm({
  "../node_modules/unenv/dist/runtime/node/perf_hooks.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_performance();
  }
});

// ../node_modules/@cloudflare/unenv-preset/dist/runtime/polyfill/performance.mjs
var init_performance2 = __esm({
  "../node_modules/@cloudflare/unenv-preset/dist/runtime/polyfill/performance.mjs"() {
    init_perf_hooks();
    globalThis.performance = performance;
    globalThis.Performance = Performance;
    globalThis.PerformanceEntry = PerformanceEntry;
    globalThis.PerformanceMark = PerformanceMark;
    globalThis.PerformanceMeasure = PerformanceMeasure;
    globalThis.PerformanceObserver = PerformanceObserver;
    globalThis.PerformanceObserverEntryList = PerformanceObserverEntryList;
    globalThis.PerformanceResourceTiming = PerformanceResourceTiming;
  }
});

// ../node_modules/unenv/dist/runtime/mock/noop.mjs
var noop_default;
var init_noop = __esm({
  "../node_modules/unenv/dist/runtime/mock/noop.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    noop_default = Object.assign(() => {
    }, { __unenv__: true });
  }
});

// ../node_modules/unenv/dist/runtime/node/console.mjs
import { Writable } from "node:stream";
var _console, _ignoreErrors, _stderr, _stdout, log, info, trace, debug, table, error, warn, createTask, clear, count, countReset, dir, dirxml, group, groupEnd, groupCollapsed, profile, profileEnd, time, timeEnd, timeLog, timeStamp, Console, _times, _stdoutErrorHandler, _stderrErrorHandler;
var init_console = __esm({
  "../node_modules/unenv/dist/runtime/node/console.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_noop();
    init_utils();
    _console = globalThis.console;
    _ignoreErrors = true;
    _stderr = new Writable();
    _stdout = new Writable();
    log = _console?.log ?? noop_default;
    info = _console?.info ?? log;
    trace = _console?.trace ?? info;
    debug = _console?.debug ?? log;
    table = _console?.table ?? log;
    error = _console?.error ?? log;
    warn = _console?.warn ?? error;
    createTask = _console?.createTask ?? /* @__PURE__ */ notImplemented("console.createTask");
    clear = _console?.clear ?? noop_default;
    count = _console?.count ?? noop_default;
    countReset = _console?.countReset ?? noop_default;
    dir = _console?.dir ?? noop_default;
    dirxml = _console?.dirxml ?? noop_default;
    group = _console?.group ?? noop_default;
    groupEnd = _console?.groupEnd ?? noop_default;
    groupCollapsed = _console?.groupCollapsed ?? noop_default;
    profile = _console?.profile ?? noop_default;
    profileEnd = _console?.profileEnd ?? noop_default;
    time = _console?.time ?? noop_default;
    timeEnd = _console?.timeEnd ?? noop_default;
    timeLog = _console?.timeLog ?? noop_default;
    timeStamp = _console?.timeStamp ?? noop_default;
    Console = _console?.Console ?? /* @__PURE__ */ notImplementedClass("console.Console");
    _times = /* @__PURE__ */ new Map();
    _stdoutErrorHandler = noop_default;
    _stderrErrorHandler = noop_default;
  }
});

// ../node_modules/@cloudflare/unenv-preset/dist/runtime/node/console.mjs
var workerdConsole, assert, clear2, context, count2, countReset2, createTask2, debug2, dir2, dirxml2, error2, group2, groupCollapsed2, groupEnd2, info2, log2, profile2, profileEnd2, table2, time2, timeEnd2, timeLog2, timeStamp2, trace2, warn2, console_default;
var init_console2 = __esm({
  "../node_modules/@cloudflare/unenv-preset/dist/runtime/node/console.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_console();
    workerdConsole = globalThis["console"];
    ({
      assert,
      clear: clear2,
      context: (
        // @ts-expect-error undocumented public API
        context
      ),
      count: count2,
      countReset: countReset2,
      createTask: (
        // @ts-expect-error undocumented public API
        createTask2
      ),
      debug: debug2,
      dir: dir2,
      dirxml: dirxml2,
      error: error2,
      group: group2,
      groupCollapsed: groupCollapsed2,
      groupEnd: groupEnd2,
      info: info2,
      log: log2,
      profile: profile2,
      profileEnd: profileEnd2,
      table: table2,
      time: time2,
      timeEnd: timeEnd2,
      timeLog: timeLog2,
      timeStamp: timeStamp2,
      trace: trace2,
      warn: warn2
    } = workerdConsole);
    Object.assign(workerdConsole, {
      Console,
      _ignoreErrors,
      _stderr,
      _stderrErrorHandler,
      _stdout,
      _stdoutErrorHandler,
      _times
    });
    console_default = workerdConsole;
  }
});

// ../node_modules/wrangler/_virtual_unenv_global_polyfill-@cloudflare-unenv-preset-node-console
var init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console = __esm({
  "../node_modules/wrangler/_virtual_unenv_global_polyfill-@cloudflare-unenv-preset-node-console"() {
    init_console2();
    globalThis.console = console_default;
  }
});

// ../node_modules/unenv/dist/runtime/node/internal/process/hrtime.mjs
var hrtime;
var init_hrtime = __esm({
  "../node_modules/unenv/dist/runtime/node/internal/process/hrtime.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    hrtime = /* @__PURE__ */ Object.assign(/* @__PURE__ */ __name(function hrtime2(startTime) {
      const now = Date.now();
      const seconds = Math.trunc(now / 1e3);
      const nanos = now % 1e3 * 1e6;
      if (startTime) {
        let diffSeconds = seconds - startTime[0];
        let diffNanos = nanos - startTime[0];
        if (diffNanos < 0) {
          diffSeconds = diffSeconds - 1;
          diffNanos = 1e9 + diffNanos;
        }
        return [diffSeconds, diffNanos];
      }
      return [seconds, nanos];
    }, "hrtime"), { bigint: /* @__PURE__ */ __name(function bigint() {
      return BigInt(Date.now() * 1e6);
    }, "bigint") });
  }
});

// ../node_modules/unenv/dist/runtime/node/internal/tty/read-stream.mjs
import { Socket } from "node:net";
var ReadStream;
var init_read_stream = __esm({
  "../node_modules/unenv/dist/runtime/node/internal/tty/read-stream.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    ReadStream = class extends Socket {
      fd;
      constructor(fd) {
        super();
        this.fd = fd;
      }
      isRaw = false;
      setRawMode(mode) {
        this.isRaw = mode;
        return this;
      }
      isTTY = false;
    };
    __name(ReadStream, "ReadStream");
  }
});

// ../node_modules/unenv/dist/runtime/node/internal/tty/write-stream.mjs
import { Socket as Socket2 } from "node:net";
var WriteStream;
var init_write_stream = __esm({
  "../node_modules/unenv/dist/runtime/node/internal/tty/write-stream.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    WriteStream = class extends Socket2 {
      fd;
      constructor(fd) {
        super();
        this.fd = fd;
      }
      clearLine(dir3, callback) {
        callback && callback();
        return false;
      }
      clearScreenDown(callback) {
        callback && callback();
        return false;
      }
      cursorTo(x, y, callback) {
        callback && typeof callback === "function" && callback();
        return false;
      }
      moveCursor(dx, dy, callback) {
        callback && callback();
        return false;
      }
      getColorDepth(env2) {
        return 1;
      }
      hasColors(count3, env2) {
        return false;
      }
      getWindowSize() {
        return [this.columns, this.rows];
      }
      columns = 80;
      rows = 24;
      isTTY = false;
    };
    __name(WriteStream, "WriteStream");
  }
});

// ../node_modules/unenv/dist/runtime/node/tty.mjs
var init_tty = __esm({
  "../node_modules/unenv/dist/runtime/node/tty.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_read_stream();
    init_write_stream();
  }
});

// ../node_modules/unenv/dist/runtime/node/internal/process/process.mjs
import { EventEmitter } from "node:events";
var Process;
var init_process = __esm({
  "../node_modules/unenv/dist/runtime/node/internal/process/process.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_tty();
    init_utils();
    Process = class extends EventEmitter {
      env;
      hrtime;
      nextTick;
      constructor(impl) {
        super();
        this.env = impl.env;
        this.hrtime = impl.hrtime;
        this.nextTick = impl.nextTick;
        for (const prop of [...Object.getOwnPropertyNames(Process.prototype), ...Object.getOwnPropertyNames(EventEmitter.prototype)]) {
          const value = this[prop];
          if (typeof value === "function") {
            this[prop] = value.bind(this);
          }
        }
      }
      emitWarning(warning, type, code) {
        console.warn(`${code ? `[${code}] ` : ""}${type ? `${type}: ` : ""}${warning}`);
      }
      emit(...args) {
        return super.emit(...args);
      }
      listeners(eventName) {
        return super.listeners(eventName);
      }
      #stdin;
      #stdout;
      #stderr;
      get stdin() {
        return this.#stdin ??= new ReadStream(0);
      }
      get stdout() {
        return this.#stdout ??= new WriteStream(1);
      }
      get stderr() {
        return this.#stderr ??= new WriteStream(2);
      }
      #cwd = "/";
      chdir(cwd2) {
        this.#cwd = cwd2;
      }
      cwd() {
        return this.#cwd;
      }
      arch = "";
      platform = "";
      argv = [];
      argv0 = "";
      execArgv = [];
      execPath = "";
      title = "";
      pid = 200;
      ppid = 100;
      get version() {
        return "";
      }
      get versions() {
        return {};
      }
      get allowedNodeEnvironmentFlags() {
        return /* @__PURE__ */ new Set();
      }
      get sourceMapsEnabled() {
        return false;
      }
      get debugPort() {
        return 0;
      }
      get throwDeprecation() {
        return false;
      }
      get traceDeprecation() {
        return false;
      }
      get features() {
        return {};
      }
      get release() {
        return {};
      }
      get connected() {
        return false;
      }
      get config() {
        return {};
      }
      get moduleLoadList() {
        return [];
      }
      constrainedMemory() {
        return 0;
      }
      availableMemory() {
        return 0;
      }
      uptime() {
        return 0;
      }
      resourceUsage() {
        return {};
      }
      ref() {
      }
      unref() {
      }
      umask() {
        throw createNotImplementedError("process.umask");
      }
      getBuiltinModule() {
        return void 0;
      }
      getActiveResourcesInfo() {
        throw createNotImplementedError("process.getActiveResourcesInfo");
      }
      exit() {
        throw createNotImplementedError("process.exit");
      }
      reallyExit() {
        throw createNotImplementedError("process.reallyExit");
      }
      kill() {
        throw createNotImplementedError("process.kill");
      }
      abort() {
        throw createNotImplementedError("process.abort");
      }
      dlopen() {
        throw createNotImplementedError("process.dlopen");
      }
      setSourceMapsEnabled() {
        throw createNotImplementedError("process.setSourceMapsEnabled");
      }
      loadEnvFile() {
        throw createNotImplementedError("process.loadEnvFile");
      }
      disconnect() {
        throw createNotImplementedError("process.disconnect");
      }
      cpuUsage() {
        throw createNotImplementedError("process.cpuUsage");
      }
      setUncaughtExceptionCaptureCallback() {
        throw createNotImplementedError("process.setUncaughtExceptionCaptureCallback");
      }
      hasUncaughtExceptionCaptureCallback() {
        throw createNotImplementedError("process.hasUncaughtExceptionCaptureCallback");
      }
      initgroups() {
        throw createNotImplementedError("process.initgroups");
      }
      openStdin() {
        throw createNotImplementedError("process.openStdin");
      }
      assert() {
        throw createNotImplementedError("process.assert");
      }
      binding() {
        throw createNotImplementedError("process.binding");
      }
      permission = { has: /* @__PURE__ */ notImplemented("process.permission.has") };
      report = {
        directory: "",
        filename: "",
        signal: "SIGUSR2",
        compact: false,
        reportOnFatalError: false,
        reportOnSignal: false,
        reportOnUncaughtException: false,
        getReport: /* @__PURE__ */ notImplemented("process.report.getReport"),
        writeReport: /* @__PURE__ */ notImplemented("process.report.writeReport")
      };
      finalization = {
        register: /* @__PURE__ */ notImplemented("process.finalization.register"),
        unregister: /* @__PURE__ */ notImplemented("process.finalization.unregister"),
        registerBeforeExit: /* @__PURE__ */ notImplemented("process.finalization.registerBeforeExit")
      };
      memoryUsage = Object.assign(() => ({
        arrayBuffers: 0,
        rss: 0,
        external: 0,
        heapTotal: 0,
        heapUsed: 0
      }), { rss: () => 0 });
      mainModule = void 0;
      domain = void 0;
      send = void 0;
      exitCode = void 0;
      channel = void 0;
      getegid = void 0;
      geteuid = void 0;
      getgid = void 0;
      getgroups = void 0;
      getuid = void 0;
      setegid = void 0;
      seteuid = void 0;
      setgid = void 0;
      setgroups = void 0;
      setuid = void 0;
      _events = void 0;
      _eventsCount = void 0;
      _exiting = void 0;
      _maxListeners = void 0;
      _debugEnd = void 0;
      _debugProcess = void 0;
      _fatalException = void 0;
      _getActiveHandles = void 0;
      _getActiveRequests = void 0;
      _kill = void 0;
      _preload_modules = void 0;
      _rawDebug = void 0;
      _startProfilerIdleNotifier = void 0;
      _stopProfilerIdleNotifier = void 0;
      _tickCallback = void 0;
      _disconnect = void 0;
      _handleQueue = void 0;
      _pendingMessage = void 0;
      _channel = void 0;
      _send = void 0;
      _linkedBinding = void 0;
    };
    __name(Process, "Process");
  }
});

// ../node_modules/@cloudflare/unenv-preset/dist/runtime/node/process.mjs
var globalProcess, getBuiltinModule, exit, platform, nextTick, unenvProcess, abort, addListener, allowedNodeEnvironmentFlags, hasUncaughtExceptionCaptureCallback, setUncaughtExceptionCaptureCallback, loadEnvFile, sourceMapsEnabled, arch, argv, argv0, chdir, config, connected, constrainedMemory, availableMemory, cpuUsage, cwd, debugPort, dlopen, disconnect, emit, emitWarning, env, eventNames, execArgv, execPath, finalization, features, getActiveResourcesInfo, getMaxListeners, hrtime3, kill, listeners, listenerCount, memoryUsage, on, off, once, pid, ppid, prependListener, prependOnceListener, rawListeners, release, removeAllListeners, removeListener, report, resourceUsage, setMaxListeners, setSourceMapsEnabled, stderr, stdin, stdout, title, throwDeprecation, traceDeprecation, umask, uptime, version, versions, domain, initgroups, moduleLoadList, reallyExit, openStdin, assert2, binding, send, exitCode, channel, getegid, geteuid, getgid, getgroups, getuid, setegid, seteuid, setgid, setgroups, setuid, permission, mainModule, _events, _eventsCount, _exiting, _maxListeners, _debugEnd, _debugProcess, _fatalException, _getActiveHandles, _getActiveRequests, _kill, _preload_modules, _rawDebug, _startProfilerIdleNotifier, _stopProfilerIdleNotifier, _tickCallback, _disconnect, _handleQueue, _pendingMessage, _channel, _send, _linkedBinding, _process, process_default;
var init_process2 = __esm({
  "../node_modules/@cloudflare/unenv-preset/dist/runtime/node/process.mjs"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    init_hrtime();
    init_process();
    globalProcess = globalThis["process"];
    getBuiltinModule = globalProcess.getBuiltinModule;
    ({ exit, platform, nextTick } = getBuiltinModule(
      "node:process"
    ));
    unenvProcess = new Process({
      env: globalProcess.env,
      hrtime,
      nextTick
    });
    ({
      abort,
      addListener,
      allowedNodeEnvironmentFlags,
      hasUncaughtExceptionCaptureCallback,
      setUncaughtExceptionCaptureCallback,
      loadEnvFile,
      sourceMapsEnabled,
      arch,
      argv,
      argv0,
      chdir,
      config,
      connected,
      constrainedMemory,
      availableMemory,
      cpuUsage,
      cwd,
      debugPort,
      dlopen,
      disconnect,
      emit,
      emitWarning,
      env,
      eventNames,
      execArgv,
      execPath,
      finalization,
      features,
      getActiveResourcesInfo,
      getMaxListeners,
      hrtime: hrtime3,
      kill,
      listeners,
      listenerCount,
      memoryUsage,
      on,
      off,
      once,
      pid,
      ppid,
      prependListener,
      prependOnceListener,
      rawListeners,
      release,
      removeAllListeners,
      removeListener,
      report,
      resourceUsage,
      setMaxListeners,
      setSourceMapsEnabled,
      stderr,
      stdin,
      stdout,
      title,
      throwDeprecation,
      traceDeprecation,
      umask,
      uptime,
      version,
      versions,
      domain,
      initgroups,
      moduleLoadList,
      reallyExit,
      openStdin,
      assert: assert2,
      binding,
      send,
      exitCode,
      channel,
      getegid,
      geteuid,
      getgid,
      getgroups,
      getuid,
      setegid,
      seteuid,
      setgid,
      setgroups,
      setuid,
      permission,
      mainModule,
      _events,
      _eventsCount,
      _exiting,
      _maxListeners,
      _debugEnd,
      _debugProcess,
      _fatalException,
      _getActiveHandles,
      _getActiveRequests,
      _kill,
      _preload_modules,
      _rawDebug,
      _startProfilerIdleNotifier,
      _stopProfilerIdleNotifier,
      _tickCallback,
      _disconnect,
      _handleQueue,
      _pendingMessage,
      _channel,
      _send,
      _linkedBinding
    } = unenvProcess);
    _process = {
      abort,
      addListener,
      allowedNodeEnvironmentFlags,
      hasUncaughtExceptionCaptureCallback,
      setUncaughtExceptionCaptureCallback,
      loadEnvFile,
      sourceMapsEnabled,
      arch,
      argv,
      argv0,
      chdir,
      config,
      connected,
      constrainedMemory,
      availableMemory,
      cpuUsage,
      cwd,
      debugPort,
      dlopen,
      disconnect,
      emit,
      emitWarning,
      env,
      eventNames,
      execArgv,
      execPath,
      exit,
      finalization,
      features,
      getBuiltinModule,
      getActiveResourcesInfo,
      getMaxListeners,
      hrtime: hrtime3,
      kill,
      listeners,
      listenerCount,
      memoryUsage,
      nextTick,
      on,
      off,
      once,
      pid,
      platform,
      ppid,
      prependListener,
      prependOnceListener,
      rawListeners,
      release,
      removeAllListeners,
      removeListener,
      report,
      resourceUsage,
      setMaxListeners,
      setSourceMapsEnabled,
      stderr,
      stdin,
      stdout,
      title,
      throwDeprecation,
      traceDeprecation,
      umask,
      uptime,
      version,
      versions,
      // @ts-expect-error old API
      domain,
      initgroups,
      moduleLoadList,
      reallyExit,
      openStdin,
      assert: assert2,
      binding,
      send,
      exitCode,
      channel,
      getegid,
      geteuid,
      getgid,
      getgroups,
      getuid,
      setegid,
      seteuid,
      setgid,
      setgroups,
      setuid,
      permission,
      mainModule,
      _events,
      _eventsCount,
      _exiting,
      _maxListeners,
      _debugEnd,
      _debugProcess,
      _fatalException,
      _getActiveHandles,
      _getActiveRequests,
      _kill,
      _preload_modules,
      _rawDebug,
      _startProfilerIdleNotifier,
      _stopProfilerIdleNotifier,
      _tickCallback,
      _disconnect,
      _handleQueue,
      _pendingMessage,
      _channel,
      _send,
      _linkedBinding
    };
    process_default = _process;
  }
});

// ../node_modules/wrangler/_virtual_unenv_global_polyfill-@cloudflare-unenv-preset-node-process
var init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process = __esm({
  "../node_modules/wrangler/_virtual_unenv_global_polyfill-@cloudflare-unenv-preset-node-process"() {
    init_process2();
    globalThis.process = process_default;
  }
});

// src/utils/auth.js
var auth_exports = {};
__export(auth_exports, {
  authenticate: () => authenticate,
  createToken: () => createToken,
  generateSessionId: () => generateSessionId,
  hasPermission: () => hasPermission,
  hashPassword: () => hashPassword,
  verifyPassword: () => verifyPassword,
  verifyToken: () => verifyToken
});
async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function verifyPassword(password, hash) {
  const computed = await hashPassword(password);
  return computed === hash;
}
function generateSessionId() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return Array.from(array).map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function createToken(payload, secret) {
  const header = btoa(JSON.stringify({ alg: "HS256", typ: "JWT" }));
  const body = btoa(JSON.stringify({ ...payload, iat: Date.now() }));
  const data = `${header}.${body}`;
  const encoder = new TextEncoder();
  const keyData = encoder.encode(secret);
  const messageData = encoder.encode(data);
  const key = await crypto.subtle.importKey(
    "raw",
    keyData,
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, messageData);
  const sigArray = Array.from(new Uint8Array(signature));
  const sig = btoa(String.fromCharCode(...sigArray));
  return `${data}.${sig}`;
}
async function verifyToken(token, secret) {
  try {
    const parts = token.split(".");
    if (parts.length !== 3)
      return null;
    const [header, body, sig] = parts;
    const data = `${header}.${body}`;
    const encoder = new TextEncoder();
    const keyData = encoder.encode(secret);
    const messageData = encoder.encode(data);
    const key = await crypto.subtle.importKey(
      "raw",
      keyData,
      { name: "HMAC", hash: "SHA-256" },
      false,
      ["verify"]
    );
    const sigBytes = Uint8Array.from(atob(sig), (c) => c.charCodeAt(0));
    const valid = await crypto.subtle.verify("HMAC", key, sigBytes, messageData);
    if (!valid)
      return null;
    const payload = JSON.parse(atob(body));
    if (payload.exp && Date.now() > payload.exp)
      return null;
    return payload;
  } catch {
    return null;
  }
}
async function authenticate(request, env2) {
  const authHeader = request.headers.get("Authorization");
  if (!authHeader || !authHeader.startsWith("Bearer "))
    return null;
  const token = authHeader.slice(7);
  const payload = await verifyToken(token, env2.JWT_SECRET || "dev-secret-change-me");
  if (!payload)
    return null;
  const user = await env2.DB.prepare(
    "SELECT id, username, email, full_name, role, is_active FROM users WHERE id = ?"
  ).bind(payload.sub).first();
  if (!user || !user.is_active)
    return null;
  return user;
}
function hasPermission(user, module, action) {
  if (user.role === "superadmin")
    return true;
  const rolePermissions = {
    admin: { level: 4 },
    manager: { level: 3 },
    spv: { level: 2 },
    viewer: { level: 1 }
  };
  const writeModules = [
    "employees",
    "contracts",
    "schedule",
    "issues",
    "one_on_one",
    "training",
    "relievers",
    "reports",
    "sop",
    "checklist",
    "forms",
    "supply_requests"
  ];
  const userLevel = rolePermissions[user.role]?.level || 0;
  if (action === "read")
    return userLevel >= 1;
  if (action === "write")
    return userLevel >= 2;
  if (action === "delete")
    return userLevel >= 3;
  if (action === "admin")
    return userLevel >= 4;
  return false;
}
var init_auth = __esm({
  "src/utils/auth.js"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    __name(hashPassword, "hashPassword");
    __name(verifyPassword, "verifyPassword");
    __name(generateSessionId, "generateSessionId");
    __name(createToken, "createToken");
    __name(verifyToken, "verifyToken");
    __name(authenticate, "authenticate");
    __name(hasPermission, "hasPermission");
  }
});

// src/utils/calendar.js
var calendar_exports = {};
__export(calendar_exports, {
  getSyncStatements: () => getSyncStatements,
  runSync: () => runSync
});
function getSyncStatements(db, module, referenceId, eventData) {
  const stmts = [];
  stmts.push(db.prepare("DELETE FROM calendar_events WHERE module = ? AND reference_id = ?").bind(module, referenceId));
  if (!eventData)
    return stmts;
  switch (module) {
    case "contracts": {
      const { empName, branchId, endDate, status } = eventData;
      if (status !== "Aktif" || !endDate)
        break;
      const intervals = [90, 60, 30, 14, 7, 3, 1, 0];
      const end = /* @__PURE__ */ new Date(endDate + "T00:00:00");
      if (!isNaN(end.getTime())) {
        intervals.forEach((days) => {
          const remDate = new Date(end.getTime());
          remDate.setDate(end.getDate() - days);
          const remStr = remDate.toISOString().slice(0, 10);
          const title2 = days === 0 ? `Hari H: Kontrak ${empName} Berakhir` : `Reminder H-${days}: Kontrak ${empName} Berakhir`;
          const priority = days <= 7 ? "critical" : days <= 30 ? "high" : "medium";
          const color = days <= 7 ? "red" : days <= 30 ? "orange" : "yellow";
          stmts.push(
            db.prepare(
              `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
            ).bind(module, referenceId, "expiry_reminder", title2, `Kontrak berakhir pada ${endDate}`, branchId, remStr, status, priority, color)
          );
        });
      }
      break;
    }
    case "schedule": {
      const { activity_type, target_date, status, pic, branch_id, notes } = eventData;
      if (!target_date)
        break;
      const color = activity_type === "Inspeksi Hygiene & Aset Bangunan" ? "blue" : activity_type === "General Cleaning" ? "green" : activity_type === "Deep Cleaning" ? "purple" : activity_type === "Fogging" ? "orange" : "gray";
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "schedule_task", activity_type, `PIC: ${pic || "-"}. Catatan: ${notes || ""}`, branch_id, target_date, status, "medium", color)
      );
      break;
    }
    case "relievers": {
      const { reliever_name, backup_date, status, branch_id, original_fc_name, reason, shift } = eventData;
      if (!backup_date)
        break;
      const desc = `Menggantikan: ${original_fc_name || "-"}. Shift: ${shift || "-"}. Alasan: ${reason || "-"}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "reliever_shift", `Reliever: ${reliever_name}`, desc, branch_id, backup_date, status, "medium", "teal")
      );
      break;
    }
    case "issues": {
      const { category, report_date, target_date, completion_date, status, branch_id, complaint, employee_name, fc_specialist } = eventData;
      const desc = `Keluhan: ${complaint || ""}. Karyawan: ${employee_name || ""}. Specialist: ${fc_specialist || ""}`;
      if (report_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, "issue_report", `Laporan Permasalahan: ${category}`, desc, branch_id, report_date, status, "high", "red")
        );
      }
      if (target_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, "issue_due", `Due Date Permasalahan: ${category}`, desc, branch_id, target_date, status, "critical", "red")
        );
      }
      if (status === "Done" && completion_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, "issue_closed", `Permasalahan Selesai: ${category}`, desc, branch_id, completion_date, status, "low", "green")
        );
      }
      break;
    }
    case "one_on_one": {
      const { employee_name, meeting_date, status, branch_id, pic, problem, solution } = eventData;
      if (!meeting_date)
        break;
      const desc = `PIC: ${pic || "-"}. Masalah: ${problem || ""}. Solusi: ${solution || ""}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "one_on_one_meeting", `One on One: ${employee_name}`, desc, branch_id, meeting_date, status, "medium", "pink")
      );
      break;
    }
    case "training": {
      const { subject, training_date, branch_id, trainer, notes, score } = eventData;
      if (!training_date)
        break;
      const desc = `Trainer: ${trainer || "-"}. Score: ${score || "-"}. Catatan: ${notes || ""}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "training_class", `Training: ${subject}`, desc, branch_id, training_date, "Done", "medium", "indigo")
      );
      break;
    }
    case "cleaning": {
      const { activity_type, activity_date, status, branch_id, notes, period } = eventData;
      if (!activity_date)
        break;
      const desc = `Periode: ${period || "-"}. Catatan: ${notes || ""}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "cleaning_report", `${activity_type}`, desc, branch_id, activity_date, status, "medium", "green")
      );
      break;
    }
    case "inspection": {
      const { inspection_date, status, branch_id, notes, period, fc_score, spv_score } = eventData;
      if (!inspection_date)
        break;
      const desc = `Score FC: ${fc_score || "-"}. Score SPV: ${spv_score || "-"}. Catatan: ${notes || ""}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "inspection_report", `Hygiene Inspection (Periode: ${period})`, desc, branch_id, inspection_date, status, "medium", "blue")
      );
      break;
    }
    case "fogging": {
      const { activity_date, status, branch_id, notes, period } = eventData;
      if (!activity_date)
        break;
      const desc = `Periode: ${period || "-"}. Catatan: ${notes || ""}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "fogging_report", `Fogging`, desc, branch_id, activity_date, status, "medium", "orange")
      );
      break;
    }
    case "basecamp": {
      const { problem, info_date, done_date, status, branch_id, pic, notes } = eventData;
      const desc = `PIC: ${pic || "-"}. Catatan: ${notes || ""}`;
      if (info_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, "basecamp_info", `Laporan Basecamp: ${problem}`, desc, branch_id, info_date, status, "medium", "purple")
        );
      }
      if (status === "Done" && done_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, "basecamp_done", `Basecamp Selesai: ${problem}`, desc, branch_id, done_date, status, "low", "green")
        );
      }
      break;
    }
    case "supply": {
      const { submitter_name, submitted_at, status, branch_id, tools_items, chemical_items } = eventData;
      if (!submitted_at)
        break;
      const items = [tools_items, chemical_items].filter(Boolean).join(", ");
      const desc = `Pengirim: ${submitter_name}. Barang: ${items || "-"}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, "supply_request", `Permintaan Barang`, desc, branch_id, submitted_at.slice(0, 10), status, "low", "brown")
      );
      break;
    }
  }
  return stmts;
}
async function runSync(db, module, referenceId, eventData) {
  const stmts = getSyncStatements(db, module, referenceId, eventData);
  if (stmts.length > 0) {
    await db.batch(stmts);
  }
}
var init_calendar = __esm({
  "src/utils/calendar.js"() {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    __name(getSyncStatements, "getSyncStatements");
    __name(runSync, "runSync");
  }
});

// ../node_modules/papaparse/papaparse.min.js
var require_papaparse_min = __commonJS({
  "../node_modules/papaparse/papaparse.min.js"(exports, module) {
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
    init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
    init_performance2();
    ((e, t) => {
      "function" == typeof define && define.amd ? define([], t) : "object" == typeof module && "undefined" != typeof exports ? module.exports = t() : e.Papa = t();
    })(exports, /* @__PURE__ */ __name(function r() {
      var n = "undefined" != typeof self ? self : "undefined" != typeof window ? window : void 0 !== n ? n : {};
      var d, s = !n.document && !!n.postMessage, a = n.IS_PAPA_WORKER || false, o = {}, h = 0, v = {};
      function P(e) {
        return 65279 === e.charCodeAt(0) ? e.slice(1) : e;
      }
      __name(P, "P");
      function u(e) {
        this._handle = null, this._finished = false, this._completed = false, this._halted = false, this._input = null, this._baseIndex = 0, this._partialLine = "", this._rowCount = 0, this._start = 0, this._nextChunk = null, this.isFirstChunk = true, this._completeResults = { data: [], errors: [], meta: {} }, function(e2) {
          var t = b(e2);
          t.chunkSize = parseInt(t.chunkSize), e2.step || e2.chunk || (t.chunkSize = null);
          this._handle = new i(t), (this._handle.streamer = this)._config = t;
        }.call(this, e), this.parseChunk = function(t, e2) {
          var i2 = parseInt(this._config.skipFirstNLines) || 0;
          if (this.isFirstChunk && 0 < i2) {
            let e3 = this._config.newline;
            e3 || (r2 = this._config.quoteChar || '"', e3 = this._handle.guessLineEndings(t, r2)), t = [...t.split(e3).slice(i2)].join(e3);
          }
          this.isFirstChunk && q(this._config.beforeFirstChunk) && void 0 !== (r2 = this._config.beforeFirstChunk(t)) && (t = r2), this.isFirstChunk = false, this._halted = false;
          var i2 = this._partialLine + t, r2 = (this._partialLine = "", this._handle.parse(i2, this._baseIndex, !this._finished));
          if (!this._handle.paused() && !this._handle.aborted()) {
            t = r2.meta.cursor, i2 = (this._finished || (this._partialLine = i2.substring(t - this._baseIndex), this._baseIndex = t), r2 && r2.data && (this._rowCount += r2.data.length), this._finished || this._config.preview && this._rowCount >= this._config.preview);
            if (a)
              n.postMessage({ results: r2, workerId: v.WORKER_ID, finished: i2 });
            else if (q(this._config.chunk) && !e2) {
              if (this._config.chunk(r2, this._handle), this._handle.paused() || this._handle.aborted())
                return void (this._halted = true);
              this._completeResults = r2 = void 0;
            }
            return this._config.step || this._config.chunk || (this._completeResults.data = this._completeResults.data.concat(r2.data), this._completeResults.errors = this._completeResults.errors.concat(r2.errors), this._completeResults.meta = r2.meta), this._completed || !i2 || !q(this._config.complete) || r2 && r2.meta.aborted || (this._config.complete(this._completeResults, this._input), this._completed = true), i2 || r2 && r2.meta.paused || this._nextChunk(), r2;
          }
          this._halted = true;
        }, this._sendError = function(e2) {
          q(this._config.error) ? this._config.error(e2) : a && this._config.error && n.postMessage({ workerId: v.WORKER_ID, error: e2, finished: false });
        };
      }
      __name(u, "u");
      function f(e) {
        var r2;
        (e = e || {}).chunkSize || (e.chunkSize = v.RemoteChunkSize), u.call(this, e), this._nextChunk = s ? function() {
          this._readChunk(), this._chunkLoaded();
        } : function() {
          this._readChunk();
        }, this.stream = function(e2) {
          this._input = e2, this._nextChunk();
        }, this._readChunk = function() {
          if (this._finished)
            this._chunkLoaded();
          else {
            if (r2 = new XMLHttpRequest(), this._config.withCredentials && (r2.withCredentials = this._config.withCredentials), s || (r2.onload = y(this._chunkLoaded, this), r2.onerror = y(this._chunkError, this)), r2.open(this._config.downloadRequestBody ? "POST" : "GET", this._input, !s), this._config.downloadRequestHeaders) {
              var e2, t = this._config.downloadRequestHeaders;
              for (e2 in t)
                r2.setRequestHeader(e2, t[e2]);
            }
            var i2;
            this._config.chunkSize && (i2 = this._start + this._config.chunkSize - 1, r2.setRequestHeader("Range", "bytes=" + this._start + "-" + i2));
            try {
              r2.send(this._config.downloadRequestBody);
            } catch (e3) {
              this._chunkError(e3.message);
            }
            s && 0 === r2.status && this._chunkError();
          }
        }, this._chunkLoaded = function() {
          4 === r2.readyState && (r2.status < 200 || 400 <= r2.status ? this._chunkError() : (this._start += this._config.chunkSize || r2.responseText.length, this._finished = !this._config.chunkSize || this._start >= ((e2) => null !== (e2 = e2.getResponseHeader("Content-Range")) ? parseInt(e2.substring(e2.lastIndexOf("/") + 1)) : -1)(r2), this.parseChunk(r2.responseText)));
        }, this._chunkError = function(e2) {
          e2 = r2.statusText || e2;
          this._sendError(new Error(e2));
        };
      }
      __name(f, "f");
      function l(e) {
        (e = e || {}).chunkSize || (e.chunkSize = v.LocalChunkSize), u.call(this, e);
        var i2, r2, n2 = "undefined" != typeof FileReader;
        this.stream = function(e2) {
          this._input = e2, r2 = e2.slice || e2.webkitSlice || e2.mozSlice, n2 ? ((i2 = new FileReader()).onload = y(this._chunkLoaded, this), i2.onerror = y(this._chunkError, this)) : i2 = new FileReaderSync(), this._nextChunk();
        }, this._nextChunk = function() {
          this._finished || this._config.preview && !(this._rowCount < this._config.preview) || this._readChunk();
        }, this._readChunk = function() {
          var e2 = this._input, t = (this._config.chunkSize && (t = Math.min(this._start + this._config.chunkSize, this._input.size), e2 = r2.call(e2, this._start, t)), i2.readAsText(e2, this._config.encoding));
          n2 || this._chunkLoaded({ target: { result: t } });
        }, this._chunkLoaded = function(e2) {
          this._start += this._config.chunkSize, this._finished = !this._config.chunkSize || this._start >= this._input.size, this.parseChunk(e2.target.result);
        }, this._chunkError = function() {
          this._sendError(i2.error);
        };
      }
      __name(l, "l");
      function c(e) {
        var i2;
        u.call(this, e = e || {}), this.stream = function(e2) {
          return i2 = e2, this._nextChunk();
        }, this._nextChunk = function() {
          var e2, t;
          if (!this._finished)
            return e2 = this._config.chunkSize, i2 = e2 ? (t = i2.substring(0, e2), i2.substring(e2)) : (t = i2, ""), this._finished = !i2, this.parseChunk(t);
        };
      }
      __name(c, "c");
      function p(e) {
        u.call(this, e = e || {});
        var t = [], i2 = true, r2 = false;
        this.pause = function() {
          u.prototype.pause.apply(this, arguments), this._input.pause();
        }, this.resume = function() {
          u.prototype.resume.apply(this, arguments), this._input.resume();
        }, this.stream = function(e2) {
          this._input = e2, this._input.on("data", this._streamData), this._input.on("end", this._streamEnd), this._input.on("error", this._streamError);
        }, this._checkIsFinished = function() {
          r2 && 1 === t.length && (this._finished = true);
        }, this._nextChunk = function() {
          this._checkIsFinished(), t.length ? this.parseChunk(t.shift()) : i2 = true;
        }, this._streamData = y(function(e2) {
          try {
            t.push("string" == typeof e2 ? e2 : e2.toString(this._config.encoding)), i2 && (i2 = false, this._checkIsFinished(), this.parseChunk(t.shift()));
          } catch (e3) {
            this._streamError(e3);
          }
        }, this), this._streamError = y(function(e2) {
          this._streamCleanUp(), this._sendError(e2);
        }, this), this._streamEnd = y(function() {
          this._streamCleanUp(), r2 = true, this._streamData("");
        }, this), this._streamCleanUp = y(function() {
          this._input.removeListener("data", this._streamData), this._input.removeListener("end", this._streamEnd), this._input.removeListener("error", this._streamError);
        }, this);
      }
      __name(p, "p");
      function i(m2) {
        var n2, s2, a2, t, o2 = Math.pow(2, 53), h2 = -o2, u2 = /^\s*-?(\d+\.?|\.\d+|\d+\.\d+)([eE][-+]?\d+)?\s*$/, d2 = /^((\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d\.\d+([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z))|(\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d([+-][0-2]\d:[0-5]\d|Z)))$/, i2 = this, r2 = 0, f2 = 0, l2 = false, e = false, c2 = [], p2 = { data: [], errors: [], meta: {} };
        function y2(e2) {
          return "greedy" === m2.skipEmptyLines ? "" === e2.join("").trim() : 1 === e2.length && 0 === e2[0].length;
        }
        __name(y2, "y");
        function g2() {
          if (p2 && a2 && (k("Delimiter", "UndetectableDelimiter", "Unable to auto-detect delimiting character; defaulted to '" + v.DefaultDelimiter + "'"), a2 = false), m2.skipEmptyLines && (p2.data = p2.data.filter(function(e3) {
            return !y2(e3);
          })), _2()) {
            let t3 = function(e3, t4) {
              e3 = P(e3), q(m2.transformHeader) && (e3 = m2.transformHeader(e3, t4)), c2.push(e3);
            };
            var t2 = t3;
            __name(t3, "t");
            if (p2)
              if (Array.isArray(p2.data[0])) {
                for (var e2 = 0; _2() && e2 < p2.data.length; e2++)
                  p2.data[e2].forEach(t3);
                p2.data.splice(0, 1);
              } else
                p2.data.forEach(t3);
          }
          function i3(e3, t3) {
            for (var i4 = m2.header ? {} : [], r4 = 0; r4 < e3.length; r4++) {
              var n3 = r4, s3 = e3[r4], s3 = ((e4, t4) => ((e5) => (m2.dynamicTypingFunction && void 0 === m2.dynamicTyping[e5] && (m2.dynamicTyping[e5] = m2.dynamicTypingFunction(e5)), true === (m2.dynamicTyping[e5] || m2.dynamicTyping)))(e4) ? "true" === t4 || "TRUE" === t4 || "false" !== t4 && "FALSE" !== t4 && (((e5) => {
                if (u2.test(e5)) {
                  e5 = parseFloat(e5);
                  if (h2 < e5 && e5 < o2)
                    return 1;
                }
              })(t4) ? parseFloat(t4) : d2.test(t4) ? new Date(t4) : "" === t4 ? null : t4) : t4)(n3 = m2.header ? r4 >= c2.length ? "__parsed_extra" : c2[r4] : n3, s3 = m2.transform ? m2.transform(s3, n3) : s3);
              "__parsed_extra" === n3 ? (i4[n3] = i4[n3] || [], i4[n3].push(s3)) : i4[n3] = s3;
            }
            return m2.header && (r4 > c2.length ? k("FieldMismatch", "TooManyFields", "Too many fields: expected " + c2.length + " fields but parsed " + r4, f2 + t3) : r4 < c2.length && k("FieldMismatch", "TooFewFields", "Too few fields: expected " + c2.length + " fields but parsed " + r4, f2 + t3)), i4;
          }
          __name(i3, "i");
          var r3;
          p2 && (m2.header || m2.dynamicTyping || m2.transform) && (r3 = 1, !p2.data.length || Array.isArray(p2.data[0]) ? (p2.data = p2.data.map(i3), r3 = p2.data.length) : p2.data = i3(p2.data, 0), m2.header && p2.meta && (p2.meta.fields = c2), f2 += r3);
        }
        __name(g2, "g");
        function _2() {
          return m2.header && 0 === c2.length;
        }
        __name(_2, "_");
        function k(e2, t2, i3, r3) {
          e2 = { type: e2, code: t2, message: i3 };
          void 0 !== r3 && (e2.row = r3), p2.errors.push(e2);
        }
        __name(k, "k");
        q(m2.step) && (t = m2.step, m2.step = function(e2) {
          p2 = e2, _2() ? g2() : (g2(), 0 !== p2.data.length && (r2 += e2.data.length, m2.preview && r2 > m2.preview ? s2.abort() : (p2.data = p2.data[0], t(p2, i2))));
        }), this.parse = function(e2, t2, i3) {
          var r3 = m2.quoteChar || '"', r3 = (m2.newline || (m2.newline = this.guessLineEndings(e2, r3)), a2 = false, m2.delimiter ? q(m2.delimiter) && (m2.delimiter = m2.delimiter(e2), p2.meta.delimiter = m2.delimiter) : ((r3 = ((e3, t3, i4, r4, n3) => {
            var s3, a3, o3, h3;
            n3 = n3 || [",", "	", "|", ";", v.RECORD_SEP, v.UNIT_SEP];
            for (var u3 = 0; u3 < n3.length; u3++) {
              for (var d3, f3 = n3[u3], l3 = 0, c3 = 0, p3 = 0, g3 = (o3 = void 0, new E({ comments: r4, delimiter: f3, newline: t3, preview: 10 }).parse(e3)), _3 = 0; _3 < g3.data.length; _3++)
                i4 && y2(g3.data[_3]) ? p3++ : (d3 = g3.data[_3].length, c3 += d3, void 0 === o3 ? o3 = d3 : 0 < d3 && (l3 += Math.abs(d3 - o3), o3 = d3));
              0 < g3.data.length && (c3 /= g3.data.length - p3), (void 0 === a3 || l3 <= a3) && (void 0 === h3 || h3 < c3) && 1.99 < c3 && (a3 = l3, s3 = f3, h3 = c3);
            }
            return { successful: !!(m2.delimiter = s3), bestDelimiter: s3 };
          })(e2, m2.newline, m2.skipEmptyLines, m2.comments, m2.delimitersToGuess)).successful ? m2.delimiter = r3.bestDelimiter : (a2 = true, m2.delimiter = v.DefaultDelimiter), p2.meta.delimiter = m2.delimiter), b(m2));
          return m2.preview && m2.header && r3.preview++, n2 = e2, s2 = new E(r3), p2 = s2.parse(n2, t2, i3), g2(), l2 ? { meta: { paused: true } } : p2 || { meta: { paused: false } };
        }, this.paused = function() {
          return l2;
        }, this.pause = function() {
          l2 = true, s2.abort(), n2 = q(m2.chunk) ? "" : n2.substring(s2.getCharIndex());
        }, this.resume = function() {
          i2.streamer._halted ? (l2 = false, i2.streamer.parseChunk(n2, true)) : setTimeout(i2.resume, 3);
        }, this.aborted = function() {
          return e;
        }, this.abort = function() {
          e = true, s2.abort(), p2.meta.aborted = true, q(m2.complete) && m2.complete(p2), n2 = "";
        }, this.guessLineEndings = function(e2, t2) {
          e2 = e2.substring(0, 1048576);
          var t2 = new RegExp(U(t2) + "([^]*?)" + U(t2), "gm"), i3 = (e2 = e2.replace(t2, "")).split("\r"), t2 = e2.split("\n"), e2 = 1 < t2.length && t2[0].length < i3[0].length;
          if (1 === i3.length || e2)
            return "\n";
          for (var r3 = 0, n3 = 0; n3 < i3.length; n3++)
            "\n" === i3[n3][0] && r3++;
          return r3 >= i3.length / 2 ? "\r\n" : "\r";
        };
      }
      __name(i, "i");
      function U(e) {
        return e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      }
      __name(U, "U");
      function E(C) {
        var S = (C = C || {}).delimiter, O = C.newline, x = C.comments, I = C.step, A = C.preview, T = C.fastMode, D = null, L = false, F = null == C.quoteChar ? '"' : C.quoteChar, j = F;
        if (void 0 !== C.escapeChar && (j = C.escapeChar), ("string" != typeof S || -1 < v.BAD_DELIMITERS.indexOf(S)) && (S = ","), x === S)
          throw new Error("Comment character same as delimiter");
        true === x ? x = "#" : ("string" != typeof x || -1 < v.BAD_DELIMITERS.indexOf(x)) && (x = false), "\n" !== O && "\r" !== O && "\r\n" !== O && (O = "\n");
        var z = 0, M = false;
        this.parse = function(i2, t, r2) {
          if ("string" != typeof i2)
            throw new Error("Input must be a string");
          var n2 = i2.length, e = S.length, s2 = O.length, a2 = x.length, o2 = q(I), h2 = [], u2 = [], d2 = [], f2 = z = 0;
          if (!i2)
            return w();
          if (T || false !== T && -1 === i2.indexOf(F)) {
            for (var l2 = i2.split(O), c2 = 0; c2 < l2.length; c2++) {
              if (d2 = l2[c2], z += d2.length, c2 !== l2.length - 1)
                z += O.length;
              else if (r2)
                return w();
              if (!x || d2.substring(0, a2) !== x) {
                if (o2) {
                  if (h2 = [], k(d2.split(S)), R(), M)
                    return w();
                } else
                  k(d2.split(S));
                if (A && A <= c2)
                  return h2 = h2.slice(0, A), w(true);
              }
            }
            return w();
          }
          for (var p2 = i2.indexOf(S, z), g2 = i2.indexOf(O, z), _2 = new RegExp(U(j) + U(F), "g"), m2 = i2.indexOf(F, z); ; )
            if (i2[z] === F)
              for (m2 = z, z++; ; ) {
                if (-1 === (m2 = i2.indexOf(F, m2 + 1)))
                  return r2 || u2.push({ type: "Quotes", code: "MissingQuotes", message: "Quoted field unterminated", row: h2.length, index: z }), E2();
                if (m2 === n2 - 1)
                  return E2(i2.substring(z, m2).replace(_2, F));
                if (F === j && i2[m2 + 1] === j)
                  m2++;
                else if (F === j || 0 === m2 || i2[m2 - 1] !== j) {
                  -1 !== p2 && p2 < m2 + 1 && (p2 = i2.indexOf(S, m2 + 1));
                  var y2 = v2(-1 === (g2 = -1 !== g2 && g2 < m2 + 1 ? i2.indexOf(O, m2 + 1) : g2) ? p2 : Math.min(p2, g2));
                  if (i2.substr(m2 + 1 + y2, e) === S) {
                    d2.push(i2.substring(z, m2).replace(_2, F)), i2[z = m2 + 1 + y2 + e] !== F && (m2 = i2.indexOf(F, z)), p2 = i2.indexOf(S, z), g2 = i2.indexOf(O, z);
                    break;
                  }
                  y2 = v2(g2);
                  if (i2.substring(m2 + 1 + y2, m2 + 1 + y2 + s2) === O) {
                    if (d2.push(i2.substring(z, m2).replace(_2, F)), b2(m2 + 1 + y2 + s2), p2 = i2.indexOf(S, z), m2 = i2.indexOf(F, z), o2 && (R(), M))
                      return w();
                    if (A && h2.length >= A)
                      return w(true);
                    break;
                  }
                  u2.push({ type: "Quotes", code: "InvalidQuotes", message: "Trailing quote on quoted field is malformed", row: h2.length, index: z }), m2++;
                }
              }
            else if (x && 0 === d2.length && i2.substring(z, z + a2) === x) {
              if (-1 === g2)
                return w();
              z = g2 + s2, g2 = i2.indexOf(O, z), p2 = i2.indexOf(S, z);
            } else if (-1 !== p2 && (p2 < g2 || -1 === g2))
              d2.push(i2.substring(z, p2)), z = p2 + e, p2 = i2.indexOf(S, z);
            else {
              if (-1 === g2)
                break;
              if (d2.push(i2.substring(z, g2)), b2(g2 + s2), o2 && (R(), M))
                return w();
              if (A && h2.length >= A)
                return w(true);
            }
          return E2();
          function k(e2) {
            h2.push(e2), f2 = z;
          }
          __name(k, "k");
          function v2(e2) {
            var t2 = 0;
            return t2 = -1 !== e2 && (e2 = i2.substring(m2 + 1, e2)) && "" === e2.trim() ? e2.length : t2;
          }
          __name(v2, "v");
          function E2(e2) {
            return r2 || (void 0 === e2 && (e2 = i2.substring(z)), d2.push(e2), z = n2, k(d2), o2 && R()), w();
          }
          __name(E2, "E");
          function b2(e2) {
            z = e2, k(d2), d2 = [], g2 = i2.indexOf(O, z);
          }
          __name(b2, "b");
          function w(e2) {
            if (C.header && !t && h2.length && !L) {
              var s3 = h2[0], a3 = /* @__PURE__ */ Object.create(null), o3 = new Set(s3);
              let n3 = false;
              for (let r3 = 0; r3 < s3.length; r3++) {
                let i3 = P(s3[r3]);
                if (a3[i3 = q(C.transformHeader) ? C.transformHeader(i3, r3) : i3]) {
                  let e3, t2 = a3[i3];
                  for (; e3 = i3 + "_" + t2, t2++, o3.has(e3); )
                    ;
                  o3.add(e3), s3[r3] = e3, a3[i3]++, n3 = true, (D = null === D ? {} : D)[e3] = i3;
                } else
                  a3[i3] = 1, s3[r3] = i3;
                o3.add(i3);
              }
              n3 && console.warn("Duplicate headers found and renamed."), L = true;
            }
            return { data: h2, errors: u2, meta: { delimiter: S, linebreak: O, aborted: M, truncated: !!e2, cursor: f2 + (t || 0), renamedHeaders: D } };
          }
          __name(w, "w");
          function R() {
            I(w()), h2 = [], u2 = [];
          }
          __name(R, "R");
        }, this.abort = function() {
          M = true;
        }, this.getCharIndex = function() {
          return z;
        };
      }
      __name(E, "E");
      function g(e) {
        var t = e.data, i2 = o[t.workerId], r2 = false;
        if (t.error)
          i2.userError(t.error, t.file);
        else if (t.results && t.results.data) {
          var n2 = { abort: function() {
            r2 = true, _(t.workerId, { data: [], errors: [], meta: { aborted: true } });
          }, pause: m, resume: m };
          if (q(i2.userStep)) {
            for (var s2 = 0; s2 < t.results.data.length && (i2.userStep({ data: t.results.data[s2], errors: t.results.errors, meta: t.results.meta }, n2), !r2); s2++)
              ;
            delete t.results;
          } else
            q(i2.userChunk) && (i2.userChunk(t.results, n2, t.file), delete t.results);
        }
        t.finished && !r2 && _(t.workerId, t.results);
      }
      __name(g, "g");
      function _(e, t) {
        var i2 = o[e];
        q(i2.userComplete) && i2.userComplete(t), i2.terminate(), delete o[e];
      }
      __name(_, "_");
      function m() {
        throw new Error("Not implemented.");
      }
      __name(m, "m");
      function b(e) {
        if ("object" != typeof e || null === e)
          return e;
        var t, i2 = Array.isArray(e) ? [] : {};
        for (t in e)
          i2[t] = b(e[t]);
        return i2;
      }
      __name(b, "b");
      function y(e, t) {
        return function() {
          e.apply(t, arguments);
        };
      }
      __name(y, "y");
      function q(e) {
        return "function" == typeof e;
      }
      __name(q, "q");
      return v.parse = function(e, t) {
        var i2 = (t = t || {}).dynamicTyping || false;
        q(i2) && (t.dynamicTypingFunction = i2, i2 = {});
        if (t.dynamicTyping = i2, t.transform = !!q(t.transform) && t.transform, !t.worker || !v.WORKERS_SUPPORTED)
          return i2 = null, v.NODE_STREAM_INPUT, "string" == typeof e ? (e = P(e), i2 = new (t.download ? f : c)(t)) : true === e.readable && q(e.read) && q(e.on) ? i2 = new p(t) : (n.File && e instanceof File || e instanceof Object) && (i2 = new l(t)), i2.stream(e);
        (i2 = (() => {
          var e2;
          return !!v.WORKERS_SUPPORTED && (e2 = (() => {
            var e3 = n.URL || n.webkitURL || null, t2 = r.toString();
            return v.BLOB_URL || (v.BLOB_URL = e3.createObjectURL(new Blob(["var global = (function() { if (typeof self !== 'undefined') { return self; } if (typeof window !== 'undefined') { return window; } if (typeof global !== 'undefined') { return global; } return {}; })(); global.IS_PAPA_WORKER=true; ", "(", t2, ")();"], { type: "text/javascript" })));
          })(), (e2 = new n.Worker(e2)).onmessage = g, e2.id = h++, o[e2.id] = e2);
        })()).userStep = t.step, i2.userChunk = t.chunk, i2.userComplete = t.complete, i2.userError = t.error, t.step = q(t.step), t.chunk = q(t.chunk), t.complete = q(t.complete), t.error = q(t.error), delete t.worker, i2.postMessage({ input: e, config: t, workerId: i2.id });
      }, v.unparse = function(e, t) {
        var s2 = false, _2 = true, m2 = ",", y2 = "\r\n", a2 = '"', o2 = a2 + a2, i2 = false, r2 = null, h2 = false, u2 = ((() => {
          if ("object" == typeof t) {
            if ("string" != typeof t.delimiter || v.BAD_DELIMITERS.filter(function(e2) {
              return -1 !== t.delimiter.indexOf(e2);
            }).length || (m2 = t.delimiter), "boolean" != typeof t.quotes && "function" != typeof t.quotes && !Array.isArray(t.quotes) || (s2 = t.quotes), "boolean" != typeof t.skipEmptyLines && "string" != typeof t.skipEmptyLines || (i2 = t.skipEmptyLines), "string" == typeof t.newline && (y2 = t.newline), "string" == typeof t.quoteChar && (a2 = t.quoteChar, o2 = a2 + a2), "boolean" == typeof t.header && (_2 = t.header), Array.isArray(t.columns)) {
              if (0 === t.columns.length)
                throw new Error("Option columns is empty");
              r2 = t.columns;
            }
            void 0 !== t.escapeChar && (o2 = t.escapeChar + a2), t.escapeFormulae instanceof RegExp ? h2 = t.escapeFormulae : "boolean" == typeof t.escapeFormulae && t.escapeFormulae && (h2 = /^[=+\-@\t\r].*$/);
          }
        })(), new RegExp(U(a2), "g"));
        "string" == typeof e && (e = JSON.parse(e));
        if (Array.isArray(e)) {
          if (!e.length || Array.isArray(e[0]))
            return n2(null, e, i2);
          if ("object" == typeof e[0])
            return n2(r2 || Object.keys(e[0]), e, i2);
        } else if ("object" == typeof e)
          return "string" == typeof e.data && (e.data = JSON.parse(e.data)), Array.isArray(e.data) && (e.fields || (e.fields = e.meta && e.meta.fields || r2), e.fields || (e.fields = Array.isArray(e.data[0]) ? e.fields : "object" == typeof e.data[0] ? Object.keys(e.data[0]) : []), Array.isArray(e.data[0]) || "object" == typeof e.data[0] || (e.data = [e.data])), n2(e.fields || [], e.data || [], i2);
        throw new Error("Unable to serialize unrecognized input");
        function n2(e2, t2, i3) {
          var r3 = "", n3 = ("string" == typeof e2 && (e2 = JSON.parse(e2)), "string" == typeof t2 && (t2 = JSON.parse(t2)), Array.isArray(e2) && 0 < e2.length), s3 = !Array.isArray(t2[0]);
          if (n3 && _2) {
            for (var a3 = 0; a3 < e2.length; a3++)
              0 < a3 && (r3 += m2), r3 += k(e2[a3], a3);
            0 < t2.length && (r3 += y2);
          }
          for (var o3 = 0; o3 < t2.length; o3++) {
            var h3 = (n3 ? e2 : t2[o3]).length, u3 = false, d2 = n3 ? 0 === Object.keys(t2[o3]).length : 0 === t2[o3].length;
            if (i3 && !n3 && (u3 = "greedy" === i3 ? "" === t2[o3].join("").trim() : 1 === t2[o3].length && 0 === t2[o3][0].length), "greedy" === i3 && n3) {
              for (var f2 = [], l2 = 0; l2 < h3; l2++) {
                var c2 = s3 ? e2[l2] : l2;
                f2.push(t2[o3][c2]);
              }
              u3 = "" === f2.join("").trim();
            }
            if (!u3) {
              for (var p2 = 0; p2 < h3; p2++) {
                0 < p2 && !d2 && (r3 += m2);
                var g2 = n3 && s3 ? e2[p2] : p2;
                r3 += k(t2[o3][g2], p2);
              }
              o3 < t2.length - 1 && (!i3 || 0 < h3 && !d2) && (r3 += y2);
            }
          }
          return r3;
        }
        __name(n2, "n");
        function k(e2, t2) {
          var i3, r3, n3;
          return null == e2 ? "" : e2.constructor === Date ? JSON.stringify(e2).slice(1, 25) : (n3 = false, h2 && "string" == typeof e2 && h2.test(e2) && (e2 = "'" + e2, n3 = true), r3 = (i3 = e2.toString()).replace(u2, o2), (n3 = n3 || true === s2 || "function" == typeof s2 && s2(e2, t2) || Array.isArray(s2) && s2[t2] || ((e3, t3) => {
            for (var i4 = 0; i4 < t3.length; i4++)
              if (-1 < e3.indexOf(t3[i4]))
                return true;
            return false;
          })(r3, v.BAD_DELIMITERS) || -1 < r3.indexOf(m2) || -1 < i3.indexOf(a2) || " " === r3.charAt(0) || " " === r3.charAt(r3.length - 1)) ? a2 + r3 + a2 : r3);
        }
        __name(k, "k");
      }, v.RECORD_SEP = String.fromCharCode(30), v.UNIT_SEP = String.fromCharCode(31), v.BYTE_ORDER_MARK = "\uFEFF", v.BAD_DELIMITERS = ["\r", "\n", '"', v.BYTE_ORDER_MARK], v.WORKERS_SUPPORTED = !s && !!n.Worker, v.NODE_STREAM_INPUT = 1, v.LocalChunkSize = 10485760, v.RemoteChunkSize = 5242880, v.DefaultDelimiter = ",", v.Parser = E, v.ParserHandle = i, v.NetworkStreamer = f, v.FileStreamer = l, v.StringStreamer = c, v.ReadableStreamStreamer = p, n.jQuery && ((d = n.jQuery).fn.parse = function(o2) {
        var i2 = o2.config || {}, h2 = [];
        return this.each(function(e2) {
          if (!("INPUT" === d(this).prop("tagName").toUpperCase() && "file" === d(this).attr("type").toLowerCase() && n.FileReader) || !this.files || 0 === this.files.length)
            return true;
          for (var t = 0; t < this.files.length; t++)
            h2.push({ file: this.files[t], inputElem: this, instanceConfig: d.extend({}, i2) });
        }), e(), this;
        function e() {
          if (0 === h2.length)
            q(o2.complete) && o2.complete();
          else {
            var e2, t, i3, r2, n2 = h2[0];
            if (q(o2.before)) {
              var s2 = o2.before(n2.file, n2.inputElem);
              if ("object" == typeof s2) {
                if ("abort" === s2.action)
                  return e2 = "AbortError", t = n2.file, i3 = n2.inputElem, r2 = s2.reason, void (q(o2.error) && o2.error({ name: e2 }, t, i3, r2));
                if ("skip" === s2.action)
                  return void u2();
                "object" == typeof s2.config && (n2.instanceConfig = d.extend(n2.instanceConfig, s2.config));
              } else if ("skip" === s2)
                return void u2();
            }
            var a2 = n2.instanceConfig.complete;
            n2.instanceConfig.complete = function(e3) {
              q(a2) && a2(e3, n2.file, n2.inputElem), u2();
            }, v.parse(n2.file, n2.instanceConfig);
          }
        }
        __name(e, "e");
        function u2() {
          h2.splice(0, 1), e();
        }
        __name(u2, "u");
      }), a && (n.onmessage = function(e) {
        e = e.data;
        void 0 === v.WORKER_ID && e && (v.WORKER_ID = e.workerId);
        "string" == typeof e.input ? n.postMessage({ workerId: v.WORKER_ID, results: v.parse(e.input, e.config), finished: true }) : (n.File && e.input instanceof File || e.input instanceof Object) && (e = v.parse(e.input, e.config)) && n.postMessage({ workerId: v.WORKER_ID, results: e, finished: true });
      }), (f.prototype = Object.create(u.prototype)).constructor = f, (l.prototype = Object.create(u.prototype)).constructor = l, (c.prototype = Object.create(c.prototype)).constructor = c, (p.prototype = Object.create(u.prototype)).constructor = p, v;
    }, "r"));
  }
});

// src/index.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();

// src/routes/auth.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();

// src/utils/response.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
function corsHeaders(origin) {
  return {
    "Access-Control-Allow-Origin": origin || "*",
    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Max-Age": "86400"
  };
}
__name(corsHeaders, "corsHeaders");
function ok(data, status = 200, origin) {
  return new Response(JSON.stringify({ success: true, data }), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) }
  });
}
__name(ok, "ok");
function paginated(data, total, page, limit, origin) {
  return new Response(JSON.stringify({
    success: true,
    data,
    pagination: { total, page, limit, pages: Math.ceil(total / limit) }
  }), {
    status: 200,
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) }
  });
}
__name(paginated, "paginated");
function error3(message, status = 400, origin) {
  return new Response(JSON.stringify({ success: false, error: message }), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders(origin) }
  });
}
__name(error3, "error");
function unauthorized(origin) {
  return error3("Unauthorized", 401, origin);
}
__name(unauthorized, "unauthorized");
function forbidden(origin) {
  return error3("Forbidden - insufficient permissions", 403, origin);
}
__name(forbidden, "forbidden");
function notFound(origin) {
  return error3("Not found", 404, origin);
}
__name(notFound, "notFound");
function options(origin) {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(origin)
  });
}
__name(options, "options");

// src/routes/auth.js
async function handleAuth(request, env2, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/auth", "");
  if (path === "/login" && request.method === "POST") {
    return handleLogin(request, env2, origin);
  }
  if (path === "/logout" && request.method === "POST") {
    return handleLogout(request, env2, origin);
  }
  if (path === "/me" && request.method === "GET") {
    return handleMe(request, env2, origin);
  }
  if (path === "/change-password" && request.method === "POST") {
    return handleChangePassword(request, env2, origin);
  }
  return error3("Not found", 404, origin);
}
__name(handleAuth, "handleAuth");
async function handleLogin(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { username, password } = body;
  if (!username || !password)
    return error3("Username and password required", 400, origin);
  const user = await env2.DB.prepare(
    "SELECT * FROM users WHERE (username = ? OR email = ?) AND is_active = 1"
  ).bind(username, username).first();
  if (!user)
    return unauthorized(origin);
  let valid = false;
  if (user.password_hash === "$2a$10$placeholder_change_this") {
    if (password === "Admin@123") {
      const hash = await hashPassword(password);
      await env2.DB.prepare("UPDATE users SET password_hash = ? WHERE id = ?").bind(hash, user.id).run();
      valid = true;
    }
  } else {
    valid = await verifyPassword(password, user.password_hash);
  }
  if (!valid)
    return unauthorized(origin);
  const secret = env2.JWT_SECRET || "dev-secret-change-me";
  const token = await createToken(
    { sub: user.id, username: user.username, role: user.role, exp: Date.now() + 8 * 60 * 60 * 1e3 },
    secret
  );
  await env2.DB.prepare("UPDATE users SET updated_at = datetime('now') WHERE id = ?").bind(user.id).run();
  return ok({
    token,
    user: { id: user.id, username: user.username, email: user.email, full_name: user.full_name, role: user.role }
  }, 200, origin);
}
__name(handleLogin, "handleLogin");
async function handleLogout(request, env2, origin) {
  return ok({ message: "Logged out" }, 200, origin);
}
__name(handleLogout, "handleLogout");
async function handleMe(request, env2, origin) {
  const { authenticate: auth } = await Promise.resolve().then(() => (init_auth(), auth_exports));
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  return ok(user, 200, origin);
}
__name(handleMe, "handleMe");
async function handleChangePassword(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { current_password, new_password } = body;
  if (!current_password || !new_password)
    return error3("Both passwords required", 400, origin);
  if (new_password.length < 6)
    return error3("New password must be at least 6 characters", 400, origin);
  const dbUser = await env2.DB.prepare("SELECT password_hash FROM users WHERE id = ?").bind(user.id).first();
  const valid = await verifyPassword(current_password, dbUser.password_hash);
  if (!valid)
    return error3("Current password incorrect", 400, origin);
  const newHash = await hashPassword(new_password);
  await env2.DB.prepare("UPDATE users SET password_hash = ?, updated_at = datetime('now') WHERE id = ?").bind(newHash, user.id).run();
  return ok({ message: "Password changed successfully" }, 200, origin);
}
__name(handleChangePassword, "handleChangePassword");

// src/routes/users.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();

// src/utils/pagination.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
function getPagination(url) {
  const params = new URL(url).searchParams;
  const page = Math.max(1, parseInt(params.get("page") || "1"));
  const limit = Math.min(100, Math.max(1, parseInt(params.get("limit") || "20")));
  const offset = (page - 1) * limit;
  return { page, limit, offset };
}
__name(getPagination, "getPagination");
function getSearchParam(url, key) {
  return new URL(url).searchParams.get(key) || "";
}
__name(getSearchParam, "getSearchParam");

// src/routes/users.js
async function handleUsers(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "users", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/users", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listUsers(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "users", "admin"))
      return forbidden(origin);
    return importUsers(request, env2, origin);
  }
  if (request.method === "POST" && path === "")
    return createUser(request, env2, user, origin);
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getUser(id, env2, origin);
    if (request.method === "PUT")
      return updateUser(id, request, env2, user, origin);
    if (request.method === "DELETE")
      return deleteUser(id, env2, user, origin);
  }
  return error3("Not found", 404, origin);
}
__name(handleUsers, "handleUsers");
async function listUsers(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const search = getSearchParam(request.url, "search");
  let query, countQuery;
  if (search) {
    query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users 
             WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ? 
             ORDER BY full_name LIMIT ? OFFSET ?`;
    countQuery = `SELECT COUNT(*) as total FROM users WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ?`;
    const s = `%${search}%`;
    const [rows2, count4] = await Promise.all([
      env2.DB.prepare(query).bind(s, s, s, limit, offset).all(),
      env2.DB.prepare(countQuery).bind(s, s, s).first()
    ]);
    return paginated(rows2.results, count4.total, page, limit, origin);
  }
  query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users ORDER BY full_name LIMIT ? OFFSET ?`;
  countQuery = `SELECT COUNT(*) as total FROM users`;
  const [rows, count3] = await Promise.all([
    env2.DB.prepare(query).bind(limit, offset).all(),
    env2.DB.prepare(countQuery).first()
  ]);
  return paginated(rows.results, count3.total, page, limit, origin);
}
__name(listUsers, "listUsers");
async function getUser(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT id, username, email, full_name, role, is_active, created_at, updated_at FROM users WHERE id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getUser, "getUser");
async function createUser(request, env2, currentUser, origin) {
  if (!hasPermission(currentUser, "users", "admin"))
    return forbidden(origin);
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { username, email, full_name, password, role } = body;
  if (!username || !email || !full_name || !password)
    return error3("username, email, full_name, password required", 400, origin);
  const existing = await env2.DB.prepare("SELECT id FROM users WHERE username = ? OR email = ?").bind(username, email).first();
  if (existing)
    return error3("Username or email already exists", 409, origin);
  const hash = await hashPassword(password);
  const result = await env2.DB.prepare(
    "INSERT INTO users (username, email, password_hash, full_name, role) VALUES (?, ?, ?, ?, ?)"
  ).bind(username, email, hash, full_name, role || "viewer").run();
  return ok({ id: result.meta.last_row_id, message: "User created" }, 201, origin);
}
__name(createUser, "createUser");
async function updateUser(id, request, env2, currentUser, origin) {
  if (!hasPermission(currentUser, "users", "admin") && currentUser.id != id)
    return forbidden(origin);
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM users WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { email, full_name, role, is_active, password } = body;
  let updates = [];
  let values = [];
  if (email) {
    updates.push("email = ?");
    values.push(email);
  }
  if (full_name) {
    updates.push("full_name = ?");
    values.push(full_name);
  }
  if (role && hasPermission(currentUser, "users", "admin")) {
    updates.push("role = ?");
    values.push(role);
  }
  if (is_active !== void 0 && hasPermission(currentUser, "users", "admin")) {
    updates.push("is_active = ?");
    values.push(is_active ? 1 : 0);
  }
  if (password && password.length >= 6) {
    const hash = await hashPassword(password);
    updates.push("password_hash = ?");
    values.push(hash);
  }
  updates.push("updated_at = datetime('now')");
  values.push(id);
  await env2.DB.prepare(`UPDATE users SET ${updates.join(", ")} WHERE id = ?`).bind(...values).run();
  return ok({ message: "User updated" }, 200, origin);
}
__name(updateUser, "updateUser");
async function deleteUser(id, env2, currentUser, origin) {
  if (!hasPermission(currentUser, "users", "admin"))
    return forbidden(origin);
  if (currentUser.id == id)
    return error3("Cannot delete yourself", 400, origin);
  const existing = await env2.DB.prepare("SELECT id FROM users WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM users WHERE id = ?").bind(id).run();
  return ok({ message: "User deleted" }, 200, origin);
}
__name(deleteUser, "deleteUser");
async function importUsers(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.username || !item.email || !item.password || !item.full_name)
      continue;
    const hash = await hashPassword(item.password);
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO users (username, email, password_hash, full_name, role) 
         VALUES (?, ?, ?, ?, ?)
         ON CONFLICT(username) DO UPDATE SET
           email = excluded.email,
           full_name = excluded.full_name,
           role = excluded.role,
           password_hash = excluded.password_hash`
      ).bind(
        item.username,
        item.email,
        hash,
        item.full_name,
        item.role || "viewer"
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} user` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importUsers, "importUsers");

// src/routes/branches.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleBranches(request, env2, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/branches", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listBranches(request, env2, origin);
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getBranch(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "branches", "write"))
        return forbidden(origin);
      return updateBranch(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "branches", "delete"))
        return forbidden(origin);
      return deleteBranch(id, env2, origin);
    }
  }
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "branches", "write"))
      return forbidden(origin);
    return importBranches(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "branches", "write"))
      return forbidden(origin);
    return createBranch(request, env2, origin);
  }
  return error3("Not found", 404, origin);
}
__name(handleBranches, "handleBranches");
async function listBranches(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const search = getSearchParam(request.url, "search");
  const all = new URL(request.url).searchParams.get("all");
  if (all === "1") {
    const rows2 = await env2.DB.prepare("SELECT * FROM branches WHERE is_active = 1 ORDER BY code").all();
    return ok(rows2.results, 200, origin);
  }
  let query, countQuery, rows, count3;
  if (search) {
    const s = `%${search}%`;
    query = "SELECT * FROM branches WHERE (name LIKE ? OR code LIKE ? OR full_name LIKE ?) ORDER BY code LIMIT ? OFFSET ?";
    countQuery = "SELECT COUNT(*) as total FROM branches WHERE name LIKE ? OR code LIKE ? OR full_name LIKE ?";
    [rows, count3] = await Promise.all([
      env2.DB.prepare(query).bind(s, s, s, limit, offset).all(),
      env2.DB.prepare(countQuery).bind(s, s, s).first()
    ]);
  } else {
    [rows, count3] = await Promise.all([
      env2.DB.prepare("SELECT * FROM branches ORDER BY code LIMIT ? OFFSET ?").bind(limit, offset).all(),
      env2.DB.prepare("SELECT COUNT(*) as total FROM branches").first()
    ]);
  }
  return paginated(rows.results, count3.total, page, limit, origin);
}
__name(listBranches, "listBranches");
async function getBranch(id, env2, origin) {
  const row = await env2.DB.prepare("SELECT * FROM branches WHERE id = ?").bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getBranch, "getBranch");
async function createBranch(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { code, name, full_name, city } = body;
  if (!code || !name || !full_name)
    return error3("code, name, full_name required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO branches (code, name, full_name, city) VALUES (?, ?, ?, ?)"
  ).bind(code, name, full_name, city || null).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(createBranch, "createBranch");
async function updateBranch(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM branches WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { name, full_name, city, is_active } = body;
  await env2.DB.prepare(
    "UPDATE branches SET name = COALESCE(?, name), full_name = COALESCE(?, full_name), city = COALESCE(?, city), is_active = COALESCE(?, is_active), updated_at = datetime('now') WHERE id = ?"
  ).bind(name || null, full_name || null, city || null, is_active !== void 0 ? is_active ? 1 : 0 : null, id).run();
  return ok({ message: "Branch updated" }, 200, origin);
}
__name(updateBranch, "updateBranch");
async function deleteBranch(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM branches WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("UPDATE branches SET is_active = 0, updated_at = datetime('now') WHERE id = ?").bind(id).run();
  return ok({ message: "Branch deactivated" }, 200, origin);
}
__name(deleteBranch, "deleteBranch");
async function importBranches(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.code || !item.name)
      continue;
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO branches (code, name, full_name, city, is_active) 
         VALUES (?, ?, ?, ?, 1)
         ON CONFLICT(code) DO UPDATE SET 
           name = excluded.name, 
           full_name = excluded.full_name, 
           city = excluded.city,
           is_active = 1,
           updated_at = datetime('now')`
      ).bind(item.code, item.name, item.full_name || item.name, item.city || null)
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} cabang` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importBranches, "importBranches");

// src/routes/employees.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleEmployees(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "employees", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/employees", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listEmployees(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "employees", "write"))
      return forbidden(origin);
    return importEmployees(request, env2, origin);
  }
  if (request.method === "DELETE" && path === "/bulk") {
    if (!hasPermission(user, "employees", "delete"))
      return forbidden(origin);
    return bulkDeleteEmployees(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "employees", "write"))
      return forbidden(origin);
    return createEmployee(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getEmployee(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "employees", "write"))
        return forbidden(origin);
      return updateEmployee(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "employees", "delete"))
        return forbidden(origin);
      return deleteEmployee(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleEmployees, "handleEmployees");
async function listEmployees(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get("search") || "";
  const branch_id = url.searchParams.get("branch_id") || "";
  const division = url.searchParams.get("division") || "";
  const status = url.searchParams.get("status") || "";
  let conditions = [];
  let values = [];
  if (search) {
    conditions.push("(e.full_name LIKE ?)");
    values.push(`%${search}%`);
  }
  if (branch_id) {
    conditions.push("e.branch_id = ?");
    values.push(branch_id);
  }
  if (division) {
    conditions.push("e.division = ?");
    values.push(division);
  }
  if (status) {
    conditions.push("e.status = ?");
    values.push(status);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const countResult = await env2.DB.prepare(
    `SELECT COUNT(*) as total FROM employees e ${where}`
  ).bind(...values).first();
  const rows = await env2.DB.prepare(
    `SELECT e.*, b.full_name as branch_name 
     FROM employees e 
     LEFT JOIN branches b ON e.branch_id = b.id 
     ${where} ORDER BY e.full_name LIMIT ? OFFSET ?`
  ).bind(...values, limit, offset).all();
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(listEmployees, "listEmployees");
async function getEmployee(id, env2, origin) {
  const row = await env2.DB.prepare(
    `SELECT e.*, b.full_name as branch_name 
     FROM employees e LEFT JOIN branches b ON e.branch_id = b.id WHERE e.id = ?`
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  const contracts = await env2.DB.prepare(
    "SELECT * FROM contracts WHERE employee_id = ? ORDER BY start_date DESC"
  ).bind(id).all();
  return ok({ ...row, contracts: contracts.results }, 200, origin);
}
__name(getEmployee, "getEmployee");
async function createEmployee(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { full_name, branch_id, division, phone, join_date, status, notes } = body;
  if (!full_name)
    return error3("full_name required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)"
  ).bind(full_name, branch_id || null, division || "FACILITY CARE", phone || null, join_date || null, status || "Aktif", notes || null).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(createEmployee, "createEmployee");
async function updateEmployee(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM employees WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { full_name, branch_id, division, phone, join_date, status, notes } = body;
  await env2.DB.prepare(
    `UPDATE employees SET 
      full_name = COALESCE(?, full_name),
      branch_id = COALESCE(?, branch_id),
      division = COALESCE(?, division),
      phone = COALESCE(?, phone),
      join_date = COALESCE(?, join_date),
      status = COALESCE(?, status),
      notes = COALESCE(?, notes),
      updated_at = datetime('now')
     WHERE id = ?`
  ).bind(
    full_name || null,
    branch_id || null,
    division || null,
    phone || null,
    join_date || null,
    status || null,
    notes || null,
    id
  ).run();
  return ok({ message: "Employee updated" }, 200, origin);
}
__name(updateEmployee, "updateEmployee");
async function deleteEmployee(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM employees WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("UPDATE employees SET status = 'Tidak Aktif', updated_at = datetime('now') WHERE id = ?").bind(id).run();
  return ok({ message: "Employee deactivated" }, 200, origin);
}
__name(deleteEmployee, "deleteEmployee");
async function importEmployees(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.full_name)
      continue;
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) 
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.full_name,
        item.branch_id || null,
        item.division || "FACILITY CARE",
        item.phone || null,
        item.join_date || null,
        item.status || "Aktif",
        item.notes || null
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} karyawan` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importEmployees, "importEmployees");
async function bulkDeleteEmployees(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { ids } = body;
  if (!Array.isArray(ids) || ids.length === 0)
    return error3("ids array required", 400, origin);
  const validIds = ids.filter((id) => Number.isInteger(Number(id))).map(Number);
  if (validIds.length === 0)
    return error3("No valid IDs provided", 400, origin);
  const placeholders = validIds.map(() => "?").join(",");
  try {
    const result = await env2.DB.prepare(
      `DELETE FROM employees WHERE id IN (${placeholders})`
    ).bind(...validIds).run();
    return ok({ message: `Berhasil menghapus ${result.meta.changes} karyawan`, deleted: result.meta.changes }, 200, origin);
  } catch (err) {
    return error3("Gagal hapus data: " + err.message, 500, origin);
  }
}
__name(bulkDeleteEmployees, "bulkDeleteEmployees");

// src/routes/contracts.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
init_calendar();
async function handleContracts(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "contracts", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/contracts", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listContracts(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "contracts", "write"))
      return forbidden(origin);
    return importContracts(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "contracts", "write"))
      return forbidden(origin);
    return createContract(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getContract(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "contracts", "write"))
        return forbidden(origin);
      return updateContract(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "contracts", "delete"))
        return forbidden(origin);
      return deleteContract(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleContracts, "handleContracts");
async function listContracts(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get("search") || "";
  const branch_id = url.searchParams.get("branch_id") || "";
  const status = url.searchParams.get("status") || "";
  const expiring_days = url.searchParams.get("expiring_days") || "";
  let conditions = ["1=1"];
  let values = [];
  if (search) {
    conditions.push("e.full_name LIKE ?");
    values.push(`%${search}%`);
  }
  if (branch_id) {
    conditions.push("c.branch_id = ?");
    values.push(branch_id);
  }
  if (status) {
    conditions.push("c.status = ?");
    values.push(status);
  }
  if (expiring_days) {
    conditions.push("c.end_date <= date('now', '+' || ? || ' days') AND c.end_date >= date('now')");
    values.push(expiring_days);
  }
  const where = "WHERE " + conditions.join(" AND ");
  const countResult = await env2.DB.prepare(
    `SELECT COUNT(*) as total FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id ${where}`
  ).bind(...values).first();
  const rows = await env2.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
      CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c 
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     ${where} ORDER BY c.end_date ASC LIMIT ? OFFSET ?`
  ).bind(...values, limit, offset).all();
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(listContracts, "listContracts");
async function getContract(id, env2, origin) {
  const row = await env2.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
      CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c 
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.id = ?`
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getContract, "getContract");
async function createContract(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes } = body;
  if (!employee_id || !start_date || !end_date)
    return error3("employee_id, start_date, end_date required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    employee_id,
    branch_id || null,
    division || "FACILITY CARE",
    start_date,
    end_date,
    contract_type || null,
    pkwt_number || null,
    status || "Aktif",
    notes || null
  ).run();
  const newId = result.meta.last_row_id;
  const emp = await env2.DB.prepare("SELECT full_name FROM employees WHERE id = ?").bind(employee_id).first();
  const empName = emp ? emp.full_name : "Karyawan";
  await runSync(env2.DB, "contracts", newId, {
    empName,
    branchId: branch_id || null,
    endDate: end_date,
    status: status || "Aktif"
  });
  return ok({ id: newId }, 201, origin);
}
__name(createContract, "createContract");
async function updateContract(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM contracts WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes } = body;
  await env2.DB.prepare(
    `UPDATE contracts SET 
      employee_id = COALESCE(?, employee_id),
      branch_id = COALESCE(?, branch_id),
      division = COALESCE(?, division),
      start_date = COALESCE(?, start_date),
      end_date = COALESCE(?, end_date),
      contract_type = COALESCE(?, contract_type),
      pkwt_number = COALESCE(?, pkwt_number),
      status = COALESCE(?, status),
      notes = COALESCE(?, notes),
      updated_at = datetime('now')
     WHERE id = ?`
  ).bind(
    employee_id || null,
    branch_id || null,
    division || null,
    start_date || null,
    end_date || null,
    contract_type || null,
    pkwt_number || null,
    status || null,
    notes || null,
    id
  ).run();
  const updatedContract = await env2.DB.prepare("SELECT employee_id, branch_id, end_date, status FROM contracts WHERE id = ?").bind(id).first();
  if (updatedContract) {
    const emp = await env2.DB.prepare("SELECT full_name FROM employees WHERE id = ?").bind(updatedContract.employee_id).first();
    const empName = emp ? emp.full_name : "Karyawan";
    await runSync(env2.DB, "contracts", id, {
      empName,
      branchId: updatedContract.branch_id,
      endDate: updatedContract.end_date,
      status: updatedContract.status
    });
  }
  return ok({ message: "Contract updated" }, 200, origin);
}
__name(updateContract, "updateContract");
async function deleteContract(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM contracts WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM contracts WHERE id = ?").bind(id).run();
  await runSync(env2.DB, "contracts", id, null);
  return ok({ message: "Contract deleted" }, 200, origin);
}
__name(deleteContract, "deleteContract");
async function importContracts(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.employee_id || !item.start_date || !item.end_date)
      continue;
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.employee_id,
        item.branch_id || null,
        item.division || "FACILITY CARE",
        item.start_date,
        item.end_date,
        item.contract_type || null,
        item.pkwt_number || null,
        item.status || "Aktif",
        item.notes || null
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} kontrak` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importContracts, "importContracts");

// src/routes/schedule.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
init_calendar();
async function handleSchedule(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "schedule", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/schedule", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listSchedule(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "schedule", "write"))
      return forbidden(origin);
    return importSchedule(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "schedule", "write"))
      return forbidden(origin);
    return createSchedule(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getSchedule(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "schedule", "write"))
        return forbidden(origin);
      return updateSchedule(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "schedule", "delete"))
        return forbidden(origin);
      return deleteSchedule(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleSchedule, "handleSchedule");
async function listSchedule(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get("branch_id") || "";
  const activity_type = url.searchParams.get("activity_type") || "";
  const period = url.searchParams.get("period") || "";
  const status = url.searchParams.get("status") || "";
  const pic = url.searchParams.get("pic") || "";
  let conditions = [];
  let values = [];
  if (branch_id) {
    conditions.push("s.branch_id = ?");
    values.push(branch_id);
  }
  if (activity_type) {
    conditions.push("s.activity_type = ?");
    values.push(activity_type);
  }
  if (period) {
    conditions.push("s.period = ?");
    values.push(period);
  }
  if (status) {
    conditions.push("s.status = ?");
    values.push(status);
  }
  if (pic) {
    conditions.push("s.pic = ?");
    values.push(pic);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM activity_schedule s ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT s.*, b.full_name as branch_name FROM activity_schedule s
       LEFT JOIN branches b ON s.branch_id = b.id
       ${where} ORDER BY s.target_date ASC, s.id DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(listSchedule, "listSchedule");
async function getSchedule(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT s.*, b.full_name as branch_name FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id WHERE s.id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getSchedule, "getSchedule");
async function createSchedule(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes } = body;
  if (!activity_type || !period)
    return error3("activity_type and period required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    branch_id || null,
    activity_type,
    period,
    pic || null,
    opening_date || null,
    target_date || null,
    completion_date || null,
    status || "Pending",
    notes || null
  ).run();
  const newId = result.meta.last_row_id;
  await runSync(env2.DB, "schedule", newId, {
    activity_type,
    period,
    pic,
    target_date,
    status: status || "Pending",
    branch_id,
    notes
  });
  return ok({ id: newId }, 201, origin);
}
__name(createSchedule, "createSchedule");
async function updateSchedule(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM activity_schedule WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes } = body;
  await env2.DB.prepare(
    `UPDATE activity_schedule SET
      branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type),
      period = COALESCE(?, period), pic = COALESCE(?, pic),
      opening_date = COALESCE(?, opening_date), target_date = COALESCE(?, target_date),
      completion_date = COALESCE(?, completion_date), status = COALESCE(?, status),
      notes = COALESCE(?, notes), updated_at = datetime('now')
     WHERE id = ?`
  ).bind(
    branch_id || null,
    activity_type || null,
    period || null,
    pic || null,
    opening_date || null,
    target_date || null,
    completion_date || null,
    status || null,
    notes || null,
    id
  ).run();
  const updated = await env2.DB.prepare("SELECT * FROM activity_schedule WHERE id = ?").bind(id).first();
  if (updated) {
    await runSync(env2.DB, "schedule", id, updated);
  }
  return ok({ message: "Schedule updated" }, 200, origin);
}
__name(updateSchedule, "updateSchedule");
async function deleteSchedule(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM activity_schedule WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM activity_schedule WHERE id = ?").bind(id).run();
  await runSync(env2.DB, "schedule", id, null);
  return ok({ message: "Schedule deleted" }, 200, origin);
}
__name(deleteSchedule, "deleteSchedule");
async function importSchedule(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.activity_type || !item.period)
      continue;
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.branch_id || null,
        item.activity_type,
        item.period,
        item.pic || null,
        item.opening_date || null,
        item.target_date || null,
        item.completion_date || null,
        item.status || "Pending",
        item.notes || null
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} jadwal kegiatan` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importSchedule, "importSchedule");

// src/routes/issues.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleIssues(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "issues", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/issues", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return listIssues(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "issues", "write"))
      return forbidden(origin);
    return importIssues(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "issues", "write"))
      return forbidden(origin);
    return createIssue(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getIssue(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "issues", "write"))
        return forbidden(origin);
      return updateIssue(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "issues", "delete"))
        return forbidden(origin);
      return deleteIssue(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleIssues, "handleIssues");
async function listIssues(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get("branch_id") || "";
  const category = url.searchParams.get("category") || "";
  const status = url.searchParams.get("status") || "";
  const search = url.searchParams.get("search") || "";
  const year = url.searchParams.get("year") || "";
  let conditions = [];
  let values = [];
  if (branch_id) {
    conditions.push("i.branch_id = ?");
    values.push(branch_id);
  }
  if (category) {
    conditions.push("i.category = ?");
    values.push(category);
  }
  if (status) {
    conditions.push("i.status = ?");
    values.push(status);
  }
  if (search) {
    conditions.push("(i.complaint LIKE ? OR i.employee_name LIKE ?)");
    values.push(`%${search}%`, `%${search}%`);
  }
  if (year) {
    conditions.push("strftime('%Y', i.report_date) = ?");
    values.push(year);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM issues i ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT i.*, b.full_name as branch_name FROM issues i
       LEFT JOIN branches b ON i.branch_id = b.id
       ${where} ORDER BY i.report_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(listIssues, "listIssues");
async function getIssue(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT i.*, b.full_name as branch_name FROM issues i LEFT JOIN branches b ON i.branch_id = b.id WHERE i.id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getIssue, "getIssue");
async function createIssue(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;
  if (!report_date || !complaint || !category)
    return error3("report_date, category, complaint required", 400, origin);
  let day_count = null;
  if (completion_date && report_date) {
    const d1 = new Date(report_date);
    const d2 = new Date(completion_date);
    day_count = Math.floor((d2 - d1) / (1e3 * 60 * 60 * 24));
  }
  const result = await env2.DB.prepare(
    "INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    report_date,
    branch_id || null,
    category,
    source || null,
    complaint,
    employee_name || null,
    fc_specialist || null,
    solution || null,
    status || "Open",
    completion_date || null,
    day_count
  ).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(createIssue, "createIssue");
async function updateIssue(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT * FROM issues WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;
  let day_count = existing.day_count;
  const rd = report_date || existing.report_date;
  const cd = completion_date !== void 0 ? completion_date : existing.completion_date;
  if (rd && cd) {
    day_count = Math.floor((new Date(cd) - new Date(rd)) / (1e3 * 60 * 60 * 24));
  }
  await env2.DB.prepare(
    `UPDATE issues SET report_date = COALESCE(?, report_date), branch_id = COALESCE(?, branch_id),
     category = COALESCE(?, category), source = COALESCE(?, source), complaint = COALESCE(?, complaint),
     employee_name = COALESCE(?, employee_name), fc_specialist = COALESCE(?, fc_specialist),
     solution = COALESCE(?, solution), status = COALESCE(?, status),
     completion_date = COALESCE(?, completion_date), day_count = ?, updated_at = datetime('now')
     WHERE id = ?`
  ).bind(
    report_date || null,
    branch_id || null,
    category || null,
    source || null,
    complaint || null,
    employee_name || null,
    fc_specialist || null,
    solution || null,
    status || null,
    completion_date || null,
    day_count,
    id
  ).run();
  return ok({ message: "Issue updated" }, 200, origin);
}
__name(updateIssue, "updateIssue");
async function deleteIssue(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM issues WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM issues WHERE id = ?").bind(id).run();
  return ok({ message: "Issue deleted" }, 200, origin);
}
__name(deleteIssue, "deleteIssue");
async function importIssues(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.report_date || !item.complaint || !item.category)
      continue;
    let day_count = null;
    if (item.completion_date && item.report_date) {
      const d1 = new Date(item.report_date);
      const d2 = new Date(item.completion_date);
      day_count = Math.floor((d2 - d1) / (1e3 * 60 * 60 * 24));
    }
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.report_date,
        item.branch_id || null,
        item.category,
        item.source || null,
        item.complaint,
        item.employee_name || null,
        item.fc_specialist || null,
        item.solution || null,
        item.status || "Open",
        item.completion_date || null,
        day_count
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} data permasalahan` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importIssues, "importIssues");

// src/routes/one_on_one.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleOneOnOne(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "one_on_one", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/one-on-one", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return list(request, env2, origin);
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "one_on_one", "write"))
      return forbidden(origin);
    return create(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getOne(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "one_on_one", "write"))
        return forbidden(origin);
      return update(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "one_on_one", "delete"))
        return forbidden(origin);
      return remove(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleOneOnOne, "handleOneOnOne");
async function list(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get("branch_id") || "";
  const status = url.searchParams.get("status") || "";
  const search = url.searchParams.get("search") || "";
  let conditions = [];
  let values = [];
  if (branch_id) {
    conditions.push("o.branch_id = ?");
    values.push(branch_id);
  }
  if (status) {
    conditions.push("o.status = ?");
    values.push(status);
  }
  if (search) {
    conditions.push("(o.employee_name LIKE ? OR o.problem LIKE ?)");
    values.push(`%${search}%`, `%${search}%`);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM one_on_one o ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT o.*, b.full_name as branch_name FROM one_on_one o
       LEFT JOIN branches b ON o.branch_id = b.id
       ${where} ORDER BY o.meeting_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(list, "list");
async function getOne(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT o.*, b.full_name as branch_name FROM one_on_one o LEFT JOIN branches b ON o.branch_id = b.id WHERE o.id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getOne, "getOne");
async function create(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  if (!meeting_date || !employee_name || !problem)
    return error3("meeting_date, employee_name, problem required", 400, origin);
  let day_count = null;
  if (completion_date && meeting_date) {
    day_count = Math.floor((new Date(completion_date) - new Date(meeting_date)) / (1e3 * 60 * 60 * 24));
  }
  const result = await env2.DB.prepare(
    "INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    meeting_date,
    branch_id || null,
    employee_name,
    pic || null,
    problem,
    solution || null,
    status || "Open",
    completion_date || null,
    day_count,
    document_link || null
  ).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(create, "create");
async function update(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT * FROM one_on_one WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  const rd = meeting_date || existing.meeting_date;
  const cd = completion_date !== void 0 ? completion_date : existing.completion_date;
  let day_count = existing.day_count;
  if (rd && cd)
    day_count = Math.floor((new Date(cd) - new Date(rd)) / (1e3 * 60 * 60 * 24));
  await env2.DB.prepare(
    `UPDATE one_on_one SET meeting_date = COALESCE(?, meeting_date), branch_id = COALESCE(?, branch_id),
     employee_name = COALESCE(?, employee_name), pic = COALESCE(?, pic), problem = COALESCE(?, problem),
     solution = COALESCE(?, solution), status = COALESCE(?, status), completion_date = COALESCE(?, completion_date),
     day_count = ?, document_link = COALESCE(?, document_link), updated_at = datetime('now') WHERE id = ?`
  ).bind(
    meeting_date || null,
    branch_id || null,
    employee_name || null,
    pic || null,
    problem || null,
    solution || null,
    status || null,
    completion_date || null,
    day_count,
    document_link || null,
    id
  ).run();
  return ok({ message: "Updated" }, 200, origin);
}
__name(update, "update");
async function remove(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM one_on_one WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM one_on_one WHERE id = ?").bind(id).run();
  return ok({ message: "Deleted" }, 200, origin);
}
__name(remove, "remove");

// src/routes/training.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleTraining(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "training", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/training", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return list2(request, env2, origin);
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "training", "write"))
      return forbidden(origin);
    return create2(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getOne2(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "training", "write"))
        return forbidden(origin);
      return update2(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "training", "delete"))
        return forbidden(origin);
      return remove2(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleTraining, "handleTraining");
async function list2(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get("search") || "";
  const year = url.searchParams.get("year") || "";
  let conditions = [];
  let values = [];
  if (search) {
    conditions.push("(t.subject LIKE ? OR t.trainer LIKE ? OR t.participants LIKE ?)");
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (year) {
    conditions.push("strftime('%Y', t.training_date) = ?");
    values.push(year);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM training t ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM training t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY t.training_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(list2, "list");
async function getOne2(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT t.*, b.full_name as branch_name FROM training t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getOne2, "getOne");
async function create2(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  if (!training_date || !subject)
    return error3("training_date and subject required", 400, origin);
  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : participants || null;
  const result = await env2.DB.prepare(
    "INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    training_date,
    batch || null,
    subject,
    participantsStr,
    branch_id || null,
    trainer || null,
    score || null,
    notes || null,
    document_link || null
  ).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(create2, "create");
async function update2(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM training WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : participants || null;
  await env2.DB.prepare(
    `UPDATE training SET training_date = COALESCE(?, training_date), batch = COALESCE(?, batch),
     subject = COALESCE(?, subject), participants = COALESCE(?, participants),
     branch_id = COALESCE(?, branch_id), trainer = COALESCE(?, trainer),
     score = COALESCE(?, score), notes = COALESCE(?, notes),
     document_link = COALESCE(?, document_link), updated_at = datetime('now') WHERE id = ?`
  ).bind(
    training_date || null,
    batch || null,
    subject || null,
    participantsStr,
    branch_id || null,
    trainer || null,
    score || null,
    notes || null,
    document_link || null,
    id
  ).run();
  return ok({ message: "Updated" }, 200, origin);
}
__name(update2, "update");
async function remove2(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM training WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM training WHERE id = ?").bind(id).run();
  return ok({ message: "Deleted" }, 200, origin);
}
__name(remove2, "remove");

// src/routes/relievers.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
init_calendar();
async function handleRelievers(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "relievers", "read"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/relievers", "");
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return list3(request, env2, origin);
  if (request.method === "POST" && path === "/import") {
    if (!hasPermission(user, "relievers", "write"))
      return forbidden(origin);
    return importRelievers(request, env2, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "relievers", "write"))
      return forbidden(origin);
    return create3(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET")
      return getOne3(id, env2, origin);
    if (request.method === "PUT") {
      if (!hasPermission(user, "relievers", "write"))
        return forbidden(origin);
      return update3(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "relievers", "delete"))
        return forbidden(origin);
      return remove3(id, env2, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleRelievers, "handleRelievers");
async function list3(request, env2, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get("branch_id") || "";
  const period = url.searchParams.get("period") || "";
  const status = url.searchParams.get("status") || "";
  const reliever_name = url.searchParams.get("reliever_name") || "";
  const search = url.searchParams.get("search") || "";
  const month = url.searchParams.get("month") || "";
  let conditions = [];
  let values = [];
  if (branch_id) {
    conditions.push("r.branch_id = ?");
    values.push(branch_id);
  }
  if (period) {
    conditions.push("r.period = ?");
    values.push(period);
  }
  if (status) {
    conditions.push("r.status = ?");
    values.push(status);
  }
  if (reliever_name) {
    conditions.push("r.reliever_name LIKE ?");
    values.push(`%${reliever_name}%`);
  }
  if (search) {
    conditions.push("(r.reliever_name LIKE ? OR r.original_fc_name LIKE ? OR r.reason LIKE ?)");
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (month) {
    conditions.push("strftime('%Y-%m', r.backup_date) = ?");
    values.push(month);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM relievers r ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT r.*, b.full_name as branch_name FROM relievers r
       LEFT JOIN branches b ON r.branch_id = b.id
       ${where} ORDER BY r.backup_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(list3, "list");
async function getOne3(id, env2, origin) {
  const row = await env2.DB.prepare(
    "SELECT r.*, b.full_name as branch_name FROM relievers r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ?"
  ).bind(id).first();
  if (!row)
    return notFound(origin);
  return ok(row, 200, origin);
}
__name(getOne3, "getOne");
async function create3(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  if (!reliever_name || !backup_date)
    return error3("reliever_name and backup_date required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(
    branch_id || null,
    original_fc_name || null,
    period || null,
    reliever_name,
    backup_date,
    completion_date || null,
    reason || null,
    shift || null,
    status || "Pending"
  ).run();
  const newId = result.meta.last_row_id;
  await runSync(env2.DB, "relievers", newId, {
    reliever_name,
    backup_date,
    status: status || "Pending",
    branch_id,
    original_fc_name,
    reason,
    shift
  });
  return ok({ id: newId }, 201, origin);
}
__name(create3, "create");
async function update3(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM relievers WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  await env2.DB.prepare(
    `UPDATE relievers SET branch_id = COALESCE(?, branch_id), original_fc_name = COALESCE(?, original_fc_name),
     period = COALESCE(?, period), reliever_name = COALESCE(?, reliever_name), backup_date = COALESCE(?, backup_date),
     completion_date = COALESCE(?, completion_date), reason = COALESCE(?, reason),
     shift = COALESCE(?, shift), status = COALESCE(?, status), updated_at = datetime('now') WHERE id = ?`
  ).bind(
    branch_id || null,
    original_fc_name || null,
    period || null,
    reliever_name || null,
    backup_date || null,
    completion_date || null,
    reason || null,
    shift || null,
    status || null,
    id
  ).run();
  const updated = await env2.DB.prepare("SELECT * FROM relievers WHERE id = ?").bind(id).first();
  if (updated) {
    await runSync(env2.DB, "relievers", id, updated);
  }
  return ok({ message: "Updated" }, 200, origin);
}
__name(update3, "update");
async function remove3(id, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id FROM relievers WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  await env2.DB.prepare("DELETE FROM relievers WHERE id = ?").bind(id).run();
  await runSync(env2.DB, "relievers", id, null);
  return ok({ message: "Deleted" }, 200, origin);
}
__name(remove3, "remove");
async function importRelievers(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  if (!Array.isArray(body))
    return error3("Payload must be an array", 400, origin);
  if (body.length === 0)
    return ok({ message: "No data to import" }, 200, origin);
  const stmts = [];
  for (const item of body) {
    if (!item.reliever_name || !item.backup_date)
      continue;
    stmts.push(
      env2.DB.prepare(
        `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.branch_id || null,
        item.original_fc_name || null,
        item.period || null,
        item.reliever_name,
        item.backup_date,
        item.completion_date || null,
        item.reason || null,
        item.shift || null,
        item.status || "Pending"
      )
    );
  }
  try {
    if (stmts.length > 0)
      await env2.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} reliefer` }, 200, origin);
  } catch (err) {
    return error3("Gagal import data: " + err.message, 500, origin);
  }
}
__name(importRelievers, "importRelievers");

// src/routes/reports.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleReports(request, env2, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/reports", "");
  if (path.startsWith("/supply"))
    return handleSupply(request, env2, origin, path.replace("/supply", ""));
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "reports", "read"))
    return forbidden(origin);
  if (path.startsWith("/inspection"))
    return handleInspection(request, env2, user, origin, path.replace("/inspection", ""));
  if (path.startsWith("/cleaning"))
    return handleCleaning(request, env2, user, origin, path.replace("/cleaning", ""));
  if (path.startsWith("/fogging"))
    return handleFogging(request, env2, user, origin, path.replace("/fogging", ""));
  if (path.startsWith("/basecamp"))
    return handleBasecamp(request, env2, user, origin, path.replace("/basecamp", ""));
  return error3("Not found", 404, origin);
}
__name(handleReports, "handleReports");
async function crudList(request, env2, origin, table3, joinClause = "", extraConditions = []) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get("branch_id") || "";
  const period = url.searchParams.get("period") || "";
  const status = url.searchParams.get("status") || "";
  const year = url.searchParams.get("year") || "";
  let conditions = [...extraConditions];
  let values = [];
  if (branch_id) {
    conditions.push("t.branch_id = ?");
    values.push(branch_id);
  }
  if (period) {
    conditions.push("t.period = ?");
    values.push(period);
  }
  if (status) {
    conditions.push("t.status = ?");
    values.push(status);
  }
  if (year) {
    const dateField = table3 === "basecamp_reports" ? "info_date" : table3 === "inspection_reports" ? "inspection_date" : "activity_date";
    conditions.push(`strftime('%Y', t.${dateField}) = ?`);
    values.push(year);
  }
  const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
  const [countResult, rows] = await Promise.all([
    env2.DB.prepare(`SELECT COUNT(*) as total FROM ${table3} t ${joinClause} ${where}`).bind(...values).first(),
    env2.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM ${table3} t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY t.id DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results, countResult.total, page, limit, origin);
}
__name(crudList, "crudList");
async function handleInspection(request, env2, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return crudList(request, env2, origin, "inspection_reports");
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "reports", "write"))
      return forbidden(origin);
    return createInspection(request, env2, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT t.*, b.full_name as branch_name FROM inspection_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      if (!hasPermission(user, "reports", "write"))
        return forbidden(origin);
      return updateInspection(id, request, env2, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "reports", "delete"))
        return forbidden(origin);
      await env2.DB.prepare("DELETE FROM inspection_reports WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleInspection, "handleInspection");
async function createInspection(request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;
  if (!period || !inspection_date)
    return error3("period and inspection_date required", 400, origin);
  const result = await env2.DB.prepare(
    "INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
  ).bind(branch_id || null, period, inspection_date, status || "Pending", fc_score || null, spv_score || null, document_link || null, notes || null).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}
__name(createInspection, "createInspection");
async function updateInspection(id, request, env2, origin) {
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const existing = await env2.DB.prepare("SELECT id FROM inspection_reports WHERE id = ?").bind(id).first();
  if (!existing)
    return notFound(origin);
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;
  await env2.DB.prepare(
    `UPDATE inspection_reports SET branch_id = COALESCE(?, branch_id), period = COALESCE(?, period),
     inspection_date = COALESCE(?, inspection_date), status = COALESCE(?, status),
     fc_score = COALESCE(?, fc_score), spv_score = COALESCE(?, spv_score),
     document_link = COALESCE(?, document_link), notes = COALESCE(?, notes),
     updated_at = datetime('now') WHERE id = ?`
  ).bind(
    branch_id || null,
    period || null,
    inspection_date || null,
    status || null,
    fc_score || null,
    spv_score || null,
    document_link || null,
    notes || null,
    id
  ).run();
  return ok({ message: "Updated" }, 200, origin);
}
__name(updateInspection, "updateInspection");
async function handleCleaning(request, env2, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return crudList(request, env2, origin, "cleaning_reports");
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "reports", "write"))
      return forbidden(origin);
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!activity_type || !period || !activity_date)
      return error3("activity_type, period, activity_date required", 400, origin);
    const result = await env2.DB.prepare("INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)").bind(branch_id || null, activity_type, period, activity_date, status || "Pending", document_link || null, notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT t.*, b.full_name as branch_name FROM cleaning_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      if (!hasPermission(user, "reports", "write"))
        return forbidden(origin);
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      await env2.DB.prepare(`UPDATE cleaning_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, activity_type || null, period || null, activity_date || null, status || null, document_link || null, notes || null, id).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "reports", "delete"))
        return forbidden(origin);
      await env2.DB.prepare("DELETE FROM cleaning_reports WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleCleaning, "handleCleaning");
async function handleFogging(request, env2, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "")
    return crudList(request, env2, origin, "fogging_reports");
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "reports", "write"))
      return forbidden(origin);
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!period)
      return error3("period required", 400, origin);
    const result = await env2.DB.prepare("INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)").bind(branch_id || null, activity_type || "Fogging", period, activity_date || null, status || "Pending", document_link || null, notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT t.*, b.full_name as branch_name FROM fogging_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      if (!hasPermission(user, "reports", "write"))
        return forbidden(origin);
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      await env2.DB.prepare(`UPDATE fogging_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, activity_type || null, period || null, activity_date || null, status || null, document_link || null, notes || null, id).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "reports", "delete"))
        return forbidden(origin);
      await env2.DB.prepare("DELETE FROM fogging_reports WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleFogging, "handleFogging");
async function handleBasecamp(request, env2, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "GET" && path === "") {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const branch_id = url.searchParams.get("branch_id") || "";
    const status = url.searchParams.get("status") || "";
    const search = url.searchParams.get("search") || "";
    let conditions = [];
    let values = [];
    if (branch_id) {
      conditions.push("t.branch_id = ?");
      values.push(branch_id);
    }
    if (status) {
      conditions.push("t.status = ?");
      values.push(status);
    }
    if (search) {
      conditions.push("(t.problem LIKE ? OR t.pic LIKE ?)");
      values.push(`%${search}%`, `%${search}%`);
    }
    const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
    const [countResult, rows] = await Promise.all([
      env2.DB.prepare(`SELECT COUNT(*) as total FROM basecamp_reports t ${where}`).bind(...values).first(),
      env2.DB.prepare(`SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id ${where} ORDER BY t.info_date DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }
  if (request.method === "POST" && path === "") {
    if (!hasPermission(user, "reports", "write"))
      return forbidden(origin);
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
    if (!problem || !info_date)
      return error3("problem and info_date required", 400, origin);
    const result = await env2.DB.prepare("INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)").bind(branch_id || null, problem, pic || null, info_date, done_date || null, status || "Open", notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      if (!hasPermission(user, "reports", "write"))
        return forbidden(origin);
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
      await env2.DB.prepare(`UPDATE basecamp_reports SET branch_id = COALESCE(?, branch_id), problem = COALESCE(?, problem), pic = COALESCE(?, pic), info_date = COALESCE(?, info_date), done_date = COALESCE(?, done_date), status = COALESCE(?, status), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, problem || null, pic || null, info_date || null, done_date || null, status || null, notes || null, id).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "reports", "delete"))
        return forbidden(origin);
      await env2.DB.prepare("DELETE FROM basecamp_reports WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleBasecamp, "handleBasecamp");
async function handleSupply(request, env2, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "POST" && path === "") {
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes } = body;
    const result = await env2.DB.prepare(
      "INSERT INTO supply_requests (submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
    ).bind(
      submitter_name || null,
      branch_id || null,
      branch_name || null,
      JSON.stringify(tools_items || []),
      tools_quantity || null,
      JSON.stringify(chemical_items || []),
      chemical_quantity || null,
      additional_notes || null
    ).run();
    return ok({ id: result.meta.last_row_id, message: "Request submitted successfully" }, 201, origin);
  }
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "reports", "read"))
    return forbidden(origin);
  if (request.method === "GET" && path === "") {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const status = url.searchParams.get("status") || "";
    let conditions = [];
    let values = [];
    if (status) {
      conditions.push("r.status = ?");
      values.push(status);
    }
    const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
    const [countResult, rows] = await Promise.all([
      env2.DB.prepare(`SELECT COUNT(*) as total FROM supply_requests r ${where}`).bind(...values).first(),
      env2.DB.prepare(`SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id ${where} ORDER BY r.submitted_at DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      if (!hasPermission(user, "reports", "write"))
        return forbidden(origin);
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { status, processed_by, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes } = body;
      const t_items = tools_items ? typeof tools_items === "string" ? tools_items.split(",").map((s) => s.trim()) : tools_items : null;
      const c_items = chemical_items ? typeof chemical_items === "string" ? chemical_items.split(",").map((s) => s.trim()) : chemical_items : null;
      await env2.DB.prepare(`
        UPDATE supply_requests SET 
          status = COALESCE(?, status), 
          processed_by = COALESCE(?, processed_by), 
          processed_at = CASE WHEN ? IS NOT NULL THEN datetime('now') ELSE processed_at END,
          submitter_name = COALESCE(?, submitter_name),
          branch_id = COALESCE(?, branch_id),
          branch_name = COALESCE(?, branch_name),
          tools_items = COALESCE(?, tools_items),
          tools_quantity = COALESCE(?, tools_quantity),
          chemical_items = COALESCE(?, chemical_items),
          chemical_quantity = COALESCE(?, chemical_quantity),
          additional_notes = COALESCE(?, additional_notes)
        WHERE id = ?
      `).bind(
        status || null,
        processed_by || null,
        status || null,
        submitter_name || null,
        branch_id || null,
        branch_name || null,
        t_items ? JSON.stringify(t_items) : null,
        tools_quantity || null,
        c_items ? JSON.stringify(c_items) : null,
        chemical_quantity || null,
        additional_notes || null,
        id
      ).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      if (!hasPermission(user, "reports", "delete"))
        return forbidden(origin);
      await env2.DB.prepare("DELETE FROM supply_requests WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleSupply, "handleSupply");

// src/routes/misc.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleMisc(request, env2, origin) {
  const url = new URL(request.url);
  const path = url.pathname;
  if (path.startsWith("/api/sop"))
    return handleTable(request, env2, origin, "sop", path.replace("/api/sop", ""));
  if (path.startsWith("/api/checklist"))
    return handleTable(request, env2, origin, "master_checklist", path.replace("/api/checklist", ""));
  if (path.startsWith("/api/forms"))
    return handleForms(request, env2, origin, path.replace("/api/forms", ""));
  if (path.startsWith("/api/pic"))
    return handlePic(request, env2, origin);
  if (path.startsWith("/api/options"))
    return handleOptions(request, env2, origin);
  return error3("Not found", 404, origin);
}
__name(handleMisc, "handleMisc");
async function handleTable(request, env2, origin, table3, path) {
  if (request.method === "GET" && path === "") {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const search = url.searchParams.get("search") || "";
    const all = url.searchParams.get("all");
    if (all === "1") {
      const rows2 = await env2.DB.prepare(`SELECT * FROM ${table3} ORDER BY name`).all();
      return ok(rows2.results, 200, origin);
    }
    let conditions = [];
    let values = [];
    if (search) {
      conditions.push("(name LIKE ? OR category LIKE ?)");
      values.push(`%${search}%`, `%${search}%`);
    }
    const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
    const [countResult, rows] = await Promise.all([
      env2.DB.prepare(`SELECT COUNT(*) as total FROM ${table3} ${where}`).bind(...values).first(),
      env2.DB.prepare(`SELECT * FROM ${table3} ${where} ORDER BY name LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "sop", "write"))
    return forbidden(origin);
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "POST" && path === "") {
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { name, category, document_link, description, version: version2, effective_date } = body;
    if (!name)
      return error3("name required", 400, origin);
    const result = await env2.DB.prepare(`INSERT INTO ${table3} (name, category, document_link, description) VALUES (?, ?, ?, ?)`).bind(name, category || null, document_link || null, description || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare(`SELECT * FROM ${table3} WHERE id = ?`).bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { name, category, document_link, description } = body;
      await env2.DB.prepare(`UPDATE ${table3} SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), updated_at = datetime('now') WHERE id = ?`).bind(name || null, category || null, document_link || null, description || null, id).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      await env2.DB.prepare(`DELETE FROM ${table3} WHERE id = ?`).bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleTable, "handleTable");
async function handleForms(request, env2, origin, path) {
  if (request.method === "GET" && path === "/public") {
    const rows = await env2.DB.prepare("SELECT * FROM master_forms WHERE is_public = 1 ORDER BY name").all();
    return ok(rows.results, 200, origin);
  }
  if (request.method === "GET" && path === "") {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const search = url.searchParams.get("search") || "";
    let conditions = [];
    let values = [];
    if (search) {
      conditions.push("(name LIKE ? OR category LIKE ?)");
      values.push(`%${search}%`, `%${search}%`);
    }
    const where = conditions.length ? "WHERE " + conditions.join(" AND ") : "";
    const [countResult, rows] = await Promise.all([
      env2.DB.prepare(`SELECT COUNT(*) as total FROM master_forms ${where}`).bind(...values).first(),
      env2.DB.prepare(`SELECT * FROM master_forms ${where} ORDER BY name LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    const user2 = await authenticate(request, env2);
    if (!user2) {
      const pubRows = await env2.DB.prepare("SELECT * FROM master_forms WHERE is_public = 1 ORDER BY name").all();
      return ok(pubRows.results, 200, origin);
    }
    return paginated(rows.results, countResult.total, page, limit, origin);
  }
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "forms", "write"))
    return forbidden(origin);
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === "POST" && path === "") {
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const { name, category, document_link, description, is_public } = body;
    if (!name)
      return error3("name required", 400, origin);
    const result = await env2.DB.prepare("INSERT INTO master_forms (name, category, document_link, description, is_public) VALUES (?, ?, ?, ?, ?)").bind(name, category || null, document_link || null, description || null, is_public ? 1 : 0).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === "GET") {
      const row = await env2.DB.prepare("SELECT * FROM master_forms WHERE id = ?").bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === "PUT") {
      let body;
      try {
        body = await request.json();
      } catch {
        return error3("Invalid JSON", 400, origin);
      }
      const { name, category, document_link, description, is_public } = body;
      await env2.DB.prepare(`UPDATE master_forms SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), is_public = COALESCE(?, is_public), updated_at = datetime('now') WHERE id = ?`).bind(name || null, category || null, document_link || null, description || null, is_public !== void 0 ? is_public ? 1 : 0 : null, id).run();
      return ok({ message: "Updated" }, 200, origin);
    }
    if (request.method === "DELETE") {
      await env2.DB.prepare("DELETE FROM master_forms WHERE id = ?").bind(id).run();
      return ok({ message: "Deleted" }, 200, origin);
    }
  }
  return error3("Not found", 404, origin);
}
__name(handleForms, "handleForms");
async function handlePic(request, env2, origin) {
  if (request.method === "GET") {
    const rows = await env2.DB.prepare("SELECT * FROM pic_list WHERE is_active = 1 ORDER BY name").all();
    return ok(rows.results, 200, origin);
  }
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "sop", "admin"))
    return forbidden(origin);
  if (request.method === "POST") {
    let body;
    try {
      body = await request.json();
    } catch {
      return error3("Invalid JSON", 400, origin);
    }
    const result = await env2.DB.prepare("INSERT OR IGNORE INTO pic_list (name, role) VALUES (?, ?)").bind(body.name, body.role || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  return error3("Not found", 404, origin);
}
__name(handlePic, "handlePic");
async function handleOptions(request, env2, origin) {
  if (request.method !== "GET")
    return error3("Method not allowed", 405, origin);
  const rows = await env2.DB.prepare("SELECT category, value FROM validation_options ORDER BY value").all();
  const data = {
    pic: [],
    activity: [],
    quarter: [],
    pkwt: []
  };
  (rows.results || []).forEach((r) => {
    if (data[r.category]) {
      data[r.category].push(r.value);
    }
  });
  return ok(data, 200, origin);
}
__name(handleOptions, "handleOptions");

// src/routes/dashboard.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleDashboard(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/dashboard", "");
  if (path === "/kpi")
    return getKPI(env2, origin);
  if (path === "/stats" || path === "")
    return getStats(env2, origin);
  if (path === "/issues-trend")
    return getIssuesTrend(env2, origin);
  if (path === "/contracts-chart")
    return getContractsChart(env2, origin);
  if (path === "/issues-summary")
    return getIssuesSummary(env2, origin);
  if (path === "/inspection-bar")
    return getInspectionBar(env2, origin);
  if (path === "/contracts-expiring")
    return getContractsExpiring(env2, origin);
  if (path === "/activity-log")
    return getActivityLog(env2, origin);
  if (path === "/calendar")
    return getCalendarEvents(request, env2, origin);
  if (path === "/activity-by-period")
    return getActivityByPeriod(env2, origin);
  if (path === "/inspection-scores")
    return getInspectionScores(request, env2, origin);
  if (path === "/notifications")
    return getNotifications(request, env2, origin);
  return ok({}, 200, origin);
}
__name(handleDashboard, "handleDashboard");
function curMonthStr() {
  const n = /* @__PURE__ */ new Date();
  return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, "0")}`;
}
__name(curMonthStr, "curMonthStr");
function prevMonthStr() {
  const n = /* @__PURE__ */ new Date();
  const pm = n.getMonth() === 0 ? `${n.getFullYear() - 1}-12` : `${n.getFullYear()}-${String(n.getMonth()).padStart(2, "0")}`;
  return pm;
}
__name(prevMonthStr, "prevMonthStr");
function monthLabels(count3 = 12) {
  const labels = [];
  const now = /* @__PURE__ */ new Date();
  for (let i = count3 - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  }
  return labels;
}
__name(monthLabels, "monthLabels");
function futureMonthLabels(count3 = 6) {
  const labels = [];
  const now = /* @__PURE__ */ new Date();
  for (let i = 0; i < count3; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  }
  return labels;
}
__name(futureMonthLabels, "futureMonthLabels");
async function getKPI(env2, origin) {
  const curM = curMonthStr();
  const prevM = prevMonthStr();
  const [
    empActive,
    empPrevMonth,
    contractActive,
    contractExp30,
    issuesOpen,
    issuesPrevOpen,
    oo1Open,
    oo1PrevOpen,
    schedPending,
    supplyPending,
    totalBranches,
    trainingCur,
    inspCur,
    cleanCur,
    fogCur,
    contractPrev
  ] = await Promise.all([
    // Uses idx_employees_status
    env2.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif'").first(),
    // Uses idx_employees_status + strftime filter
    env2.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevM).first(),
    // Uses idx_contracts_status_end
    env2.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date>=date('now')").first(),
    // Uses idx_contracts_status_end (covering index)
    env2.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')").first(),
    // Uses idx_issues_status
    env2.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done' AND strftime('%Y-%m',report_date)=?").bind(prevM).first(),
    // Uses idx_oo1_status
    env2.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done' AND strftime('%Y-%m',meeting_date)=?").bind(prevM).first(),
    // Uses idx_schedule_status
    env2.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status='Pending'").first(),
    // Uses idx_supply_status
    env2.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status='Pending'").first(),
    // branches is tiny table
    env2.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1").first(),
    // Uses idx_training_date
    env2.DB.prepare("SELECT COUNT(*) c FROM training WHERE strftime('%Y-%m',training_date)=?").bind(curM).first(),
    // Uses idx_inspection_date
    env2.DB.prepare("SELECT COUNT(*) c FROM inspection_reports WHERE strftime('%Y-%m',inspection_date)=?").bind(curM).first(),
    // Uses idx_cleaning_date
    env2.DB.prepare("SELECT COUNT(*) c FROM cleaning_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curM).first(),
    // Uses idx_fogging_date
    env2.DB.prepare("SELECT COUNT(*) c FROM fogging_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curM).first(),
    // Uses idx_contracts_created
    env2.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevM).first()
  ]);
  return ok({
    employees: { current: empActive?.c || 0, prev: empPrevMonth?.c || 0 },
    contracts: { current: contractActive?.c || 0, prev: contractPrev?.c || 0 },
    expiring30: { current: contractExp30?.c || 0 },
    issues: { current: issuesOpen?.c || 0, prev: issuesPrevOpen?.c || 0 },
    one_on_one: { current: oo1Open?.c || 0, prev: oo1PrevOpen?.c || 0 },
    schedule: { current: schedPending?.c || 0 },
    supply: { current: supplyPending?.c || 0 },
    branches: { current: totalBranches?.c || 0 },
    training_month: { current: trainingCur?.c || 0 },
    inspection_month: { current: inspCur?.c || 0 },
    cleaning_month: { current: cleanCur?.c || 0 },
    fogging_month: { current: fogCur?.c || 0 }
  }, 200, origin);
}
__name(getKPI, "getKPI");
async function getStats(env2, origin) {
  const [
    totalEmployees,
    totalBranches,
    activeContracts,
    expiringContracts,
    openIssues,
    openOneOnOne,
    pendingSchedule,
    pendingSupply,
    doneIssues,
    totalTraining
  ] = await Promise.all([
    env2.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date >= date('now')").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status != 'Done'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status != 'Done'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status = 'Pending'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status = 'Pending'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status = 'Done'").first(),
    env2.DB.prepare("SELECT COUNT(*) c FROM training").first()
  ]);
  const recentIssues = await env2.DB.prepare(
    `SELECT i.id, i.report_date, i.category, i.complaint, i.status,
     b.full_name as branch_name
     FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
     ORDER BY i.created_at DESC LIMIT 5`
  ).all();
  const expiringList = await env2.DB.prepare(
    `SELECT c.id, c.end_date, c.status, e.full_name as emp_name,
     b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c LEFT JOIN branches b ON c.branch_id = b.id
     LEFT JOIN employees e ON c.employee_id = e.id
     WHERE c.status='Aktif' AND c.end_date BETWEEN date('now') AND date('now','+30 days')
     ORDER BY c.end_date ASC LIMIT 10`
  ).all();
  const upcomingSchedule = await env2.DB.prepare(
    `SELECT s.id, s.activity_type, s.target_date, s.status,
     b.full_name as branch_name
     FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id
     WHERE s.status != 'Done' AND s.target_date >= date('now')
     ORDER BY s.target_date ASC LIMIT 7`
  ).all();
  return ok({
    stats: {
      total_employees: totalEmployees?.c || 0,
      active_contracts: activeContracts?.c || 0,
      expiring_contracts: expiringContracts?.c || 0,
      open_issues: openIssues?.c || 0,
      open_one_on_one: openOneOnOne?.c || 0,
      pending_schedule: pendingSchedule?.c || 0,
      pending_supply_requests: pendingSupply?.c || 0,
      total_branches: totalBranches?.c || 0,
      resolved_issues: doneIssues?.c || 0,
      total_training: totalTraining?.c || 0
    },
    recent_issues: recentIssues.results || [],
    expiring_contracts: expiringList.results || [],
    upcoming_schedule: upcomingSchedule.results || []
  }, 200, origin);
}
__name(getStats, "getStats");
async function getIssuesTrend(env2, origin) {
  const since12m = /* @__PURE__ */ new Date();
  since12m.setMonth(since12m.getMonth() - 11);
  const since = `${since12m.getFullYear()}-${String(since12m.getMonth() + 1).padStart(2, "0")}-01`;
  const [openRows, closedRows] = await Promise.all([
    env2.DB.prepare(
      `SELECT strftime('%Y-%m', report_date) m, COUNT(*) c
       FROM issues
       WHERE report_date >= ? AND status != 'Done'
       GROUP BY m ORDER BY m`
    ).bind(since).all(),
    env2.DB.prepare(
      `SELECT strftime('%Y-%m', completion_date) m, COUNT(*) c
       FROM issues
       WHERE completion_date >= ? AND status = 'Done'
       GROUP BY m ORDER BY m`
    ).bind(since).all()
  ]);
  const labels = monthLabels(12);
  const openMap = Object.fromEntries((openRows.results || []).map((r) => [r.m, r.c]));
  const closedMap = Object.fromEntries((closedRows.results || []).map((r) => [r.m, r.c]));
  return ok({
    labels,
    open: labels.map((l) => openMap[l] || 0),
    closed: labels.map((l) => closedMap[l] || 0)
  }, 200, origin);
}
__name(getIssuesTrend, "getIssuesTrend");
async function getContractsChart(env2, origin) {
  const rows = await env2.DB.prepare(
    `SELECT strftime('%Y-%m', end_date) m, COUNT(*) c
     FROM contracts
     WHERE status='Aktif' AND end_date >= date('now') AND end_date <= date('now','+6 months')
     GROUP BY m ORDER BY m`
  ).all();
  const labels = futureMonthLabels(6);
  const map = Object.fromEntries((rows.results || []).map((r) => [r.m, r.c]));
  return ok({ labels, counts: labels.map((l) => map[l] || 0) }, 200, origin);
}
__name(getContractsChart, "getContractsChart");
async function getIssuesSummary(env2, origin) {
  const [byCategory, byStatus, byBranch] = await Promise.all([
    env2.DB.prepare(
      `SELECT COALESCE(category,'Lainnya') category, COUNT(*) c
       FROM issues GROUP BY category ORDER BY c DESC LIMIT 10`
    ).all(),
    env2.DB.prepare(
      "SELECT COALESCE(status,'Tidak Diketahui') status, COUNT(*) c FROM issues GROUP BY status"
    ).all(),
    env2.DB.prepare(
      `SELECT b.full_name branch_name, COUNT(*) c
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       GROUP BY i.branch_id ORDER BY c DESC LIMIT 10`
    ).all()
  ]);
  return ok({
    by_category: (byCategory.results || []).map((r) => ({ category: r.category || "Lainnya", count: r.c || 0 })),
    by_status: (byStatus.results || []).map((r) => ({ status: r.status || "\u2014", count: r.c || 0 })),
    by_branch: (byBranch.results || []).map((r) => ({ branch_name: r.branch_name || "Tidak diketahui", count: r.c || 0 }))
  }, 200, origin);
}
__name(getIssuesSummary, "getIssuesSummary");
async function getInspectionBar(env2, origin) {
  const since = (() => {
    const d = /* @__PURE__ */ new Date();
    d.setMonth(d.getMonth() - 5);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`;
  })();
  const rows = await env2.DB.prepare(
    `SELECT b.full_name branch_name,
     ROUND(AVG(CAST(r.fc_score AS REAL)),1) avg_fc,
     ROUND(AVG(CAST(r.spv_score AS REAL)),1) avg_spv,
     COUNT(*) total
     FROM inspection_reports r
     LEFT JOIN branches b ON r.branch_id=b.id
     WHERE r.inspection_date >= ?
     GROUP BY r.branch_id
     ORDER BY avg_fc DESC LIMIT 15`
  ).bind(since).all();
  const results = rows.results || [];
  return ok({
    labels: results.map((r) => r.branch_name || "N/A"),
    fc: results.map((r) => r.avg_fc || 0),
    spv: results.map((r) => r.avg_spv || 0),
    totals: results.map((r) => r.total || 0)
  }, 200, origin);
}
__name(getInspectionBar, "getInspectionBar");
async function getContractsExpiring(env2, origin) {
  const rows = await env2.DB.prepare(
    `SELECT c.id, c.end_date, c.status, c.contract_type, c.pkwt_number,
     COALESCE(e.full_name, '\u2014') emp_name,
     b.full_name branch_name,
     CAST(julianday(c.end_date)-julianday('now') AS INTEGER) days_remaining
     FROM contracts c
     LEFT JOIN employees e ON c.employee_id=e.id
     LEFT JOIN branches b ON c.branch_id=b.id
     WHERE c.status='Aktif' AND c.end_date >= date('now')
     ORDER BY c.end_date ASC LIMIT 50`
  ).all();
  return ok(rows.results || [], 200, origin);
}
__name(getContractsExpiring, "getContractsExpiring");
async function getActivityLog(env2, origin) {
  const [issues, contracts, employees, oo1, training, supply, relievers, inspection] = await Promise.all([
    env2.DB.prepare(
      `SELECT 'issue' type, COALESCE(complaint,'\u2014') label, created_at,
         branch_id, category info
         FROM issues ORDER BY created_at DESC LIMIT 5`
    ).all(),
    env2.DB.prepare(
      `SELECT 'contract' type, COALESCE((SELECT full_name FROM employees WHERE id=contracts.employee_id),'Karyawan') label, created_at,
         branch_id, contract_type info
         FROM contracts ORDER BY created_at DESC LIMIT 5`
    ).all(),
    env2.DB.prepare(
      `SELECT 'employee' type, full_name label, created_at,
         branch_id, status info
         FROM employees ORDER BY created_at DESC LIMIT 5`
    ).all(),
    env2.DB.prepare(
      `SELECT 'one_on_one' type, COALESCE(employee_name,'\u2014') label, created_at,
         branch_id, status info
         FROM one_on_one ORDER BY created_at DESC LIMIT 4`
    ).all(),
    env2.DB.prepare(
      `SELECT 'training' type, COALESCE(subject,'Training') label, created_at,
         NULL branch_id, trainer info
         FROM training ORDER BY created_at DESC LIMIT 4`
    ).all(),
    env2.DB.prepare(
      `SELECT 'supply' type, COALESCE(submitter_name,'Pemohon') label,
         submitted_at created_at, NULL branch_id, status info
         FROM supply_requests ORDER BY submitted_at DESC LIMIT 4`
    ).all(),
    env2.DB.prepare(
      `SELECT 'reliever' type, COALESCE(reliever_name,'Reliefer') label, created_at,
         branch_id, reason info
         FROM relievers ORDER BY created_at DESC LIMIT 3`
    ).all(),
    env2.DB.prepare(
      `SELECT 'inspection' type, period label, created_at,
         branch_id, status info
         FROM inspection_reports ORDER BY created_at DESC LIMIT 3`
    ).all()
  ]);
  const bRows = await env2.DB.prepare("SELECT id, full_name FROM branches").all();
  const branchMap = Object.fromEntries((bRows.results || []).map((b) => [b.id, b.full_name]));
  const all = [
    ...issues.results || [],
    ...contracts.results || [],
    ...employees.results || [],
    ...oo1.results || [],
    ...training.results || [],
    ...supply.results || [],
    ...relievers.results || [],
    ...inspection.results || []
  ].map((r) => ({
    ...r,
    branch: r.branch_id ? branchMap[r.branch_id] || null : null,
    label: r.label && r.label !== "[object Object]" ? String(r.label).slice(0, 80) : "\u2014"
  })).sort((a, b) => (b.created_at || "").localeCompare(a.created_at || "")).slice(0, 20);
  return ok(all, 200, origin);
}
__name(getActivityLog, "getActivityLog");
async function getCalendarEvents(request, env2, origin) {
  const url = new URL(request.url);
  const curM = url.searchParams.get("month") || curMonthStr();
  const bind = [curM];
  const dateFilter = /* @__PURE__ */ __name((col) => `strftime('%Y-%m',${col})=?`, "dateFilter");
  const [
    sched,
    issR,
    relR,
    trainR,
    oneR,
    cleanR,
    inspR,
    fogR,
    baseR,
    supplyR,
    contrList
  ] = await Promise.all([
    env2.DB.prepare(
      `SELECT s.id,'schedule' type,s.target_date event_date,s.activity_type title,
       s.status,s.pic,b.full_name branch_name,
       CASE s.activity_type
         WHEN 'Inspeksi Hygiene & Aset Bangunan' THEN 'blue'
         WHEN 'General Cleaning' THEN 'green'
         WHEN 'Deep Cleaning' THEN 'purple'
         WHEN 'Fogging' THEN 'orange'
         ELSE 'gray' END color
       FROM activity_schedule s LEFT JOIN branches b ON s.branch_id=b.id
       WHERE ${dateFilter("s.target_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT i.id,'issue' type,i.report_date event_date,i.category title,
       i.status,b.full_name branch_name,'red' color
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       WHERE ${dateFilter("i.report_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT r.id,'reliever' type,r.backup_date event_date,r.reliever_name title,
       r.status,b.full_name branch_name,'teal' color
       FROM relievers r LEFT JOIN branches b ON r.branch_id=b.id
       WHERE ${dateFilter("r.backup_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT t.id,'training' type,t.training_date event_date,t.subject title,
       'Done' status,'' branch_name,'indigo' color
       FROM training t WHERE ${dateFilter("t.training_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT o.id,'one_on_one' type,o.meeting_date event_date,'One on One: ' || o.employee_name title,
       o.status,b.full_name branch_name,'pink' color
       FROM one_on_one o LEFT JOIN branches b ON o.branch_id=b.id
       WHERE ${dateFilter("o.meeting_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT c.id,'cleaning' type,c.activity_date event_date,c.activity_type || ': ' || b.full_name title,
       c.status,b.full_name branch_name,'green' color
       FROM cleaning_reports c LEFT JOIN branches b ON c.branch_id=b.id
       WHERE ${dateFilter("c.activity_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT r.id,'inspection' type,r.inspection_date event_date,'Inspeksi: ' || b.full_name title,
       r.status,b.full_name branch_name,'blue' color
       FROM inspection_reports r LEFT JOIN branches b ON r.branch_id=b.id
       WHERE ${dateFilter("r.inspection_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT f.id,'fogging' type,f.activity_date event_date,'Fogging: ' || b.full_name title,
       f.status,b.full_name branch_name,'orange' color
       FROM fogging_reports f LEFT JOIN branches b ON f.branch_id=b.id
       WHERE ${dateFilter("f.activity_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT bp.id,'basecamp' type,bp.info_date event_date,'Basecamp: ' || bp.problem title,
       bp.status,b.full_name branch_name,'purple' color
       FROM basecamp_reports bp LEFT JOIN branches b ON bp.branch_id=b.id
       WHERE ${dateFilter("bp.info_date")}`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT s.id,'supply' type,strftime('%Y-%m-%d',s.submitted_at) event_date,'Permintaan: ' || s.submitter_name title,
       s.status,COALESCE(b.full_name,s.branch_name) branch_name,'brown' color
       FROM supply_requests s LEFT JOIN branches b ON s.branch_id=b.id
       WHERE strftime('%Y-%m',s.submitted_at)=?`
    ).bind(...bind).all(),
    env2.DB.prepare(
      `SELECT c.id, c.end_date, COALESCE(e.full_name,'?') emp_name
       FROM contracts c LEFT JOIN employees e ON c.employee_id=e.id
       WHERE c.status='Aktif'`
    ).all()
  ]);
  const contractEvents = [];
  const intervals = [90, 60, 30, 14, 7, 1, 0];
  (contrList.results || []).forEach((c) => {
    if (!c.end_date)
      return;
    const end = /* @__PURE__ */ new Date(c.end_date + "T00:00:00");
    if (isNaN(end.getTime()))
      return;
    intervals.forEach((days) => {
      const remDate = new Date(end.getTime());
      remDate.setDate(end.getDate() - days);
      const remStr = remDate.toISOString().slice(0, 10);
      if (remStr.startsWith(curM)) {
        const daysRemaining = Math.ceil((end.getTime() - Date.now()) / 864e5);
        const title2 = days === 0 ? `Hari H: Kontrak ${c.emp_name} Berakhir` : `Reminder H-${days}: Kontrak ${c.emp_name} Berakhir`;
        contractEvents.push({
          id: c.id,
          type: "contract_expiry",
          event_date: remStr,
          title: title2,
          status: "Aktif",
          branch_name: "",
          days_remaining: daysRemaining,
          color: days <= 7 ? "red" : days <= 30 ? "orange" : "yellow"
        });
      }
    });
  });
  const events = [
    ...sched.results || [],
    ...issR.results || [],
    ...relR.results || [],
    ...trainR.results || [],
    ...oneR.results || [],
    ...cleanR.results || [],
    ...inspR.results || [],
    ...fogR.results || [],
    ...baseR.results || [],
    ...supplyR.results || [],
    ...contractEvents
  ].sort((a, b) => (a.event_date || "").localeCompare(b.event_date || ""));
  return ok(events, 200, origin);
}
__name(getCalendarEvents, "getCalendarEvents");
async function getActivityByPeriod(env2, origin) {
  const rows = await env2.DB.prepare(
    `SELECT activity_type,period,COUNT(*) total,
     SUM(CASE WHEN status='Done' THEN 1 ELSE 0 END) done
     FROM activity_schedule GROUP BY activity_type,period`
  ).all();
  return ok(rows.results || [], 200, origin);
}
__name(getActivityByPeriod, "getActivityByPeriod");
async function getInspectionScores(request, env2, origin) {
  const url = new URL(request.url);
  const period = url.searchParams.get("period") || "";
  const where = period ? "WHERE r.period=?" : "";
  const bind = period ? [period] : [];
  const rows = await env2.DB.prepare(
    `SELECT r.id,r.inspection_date,r.period,r.fc_score,r.spv_score,r.status,
     b.full_name branch_name
     FROM inspection_reports r LEFT JOIN branches b ON r.branch_id=b.id
     ${where} ORDER BY r.inspection_date ASC LIMIT 500`
  ).bind(...bind).all();
  return ok(rows.results || [], 200, origin);
}
__name(getInspectionScores, "getInspectionScores");
async function getNotifications(request, env2, origin) {
  const [contracts, pendingSched, openIssues, pendingSupply] = await Promise.all([
    env2.DB.prepare(
      `SELECT c.id, c.end_date, COALESCE(e.full_name, '?') as emp_name
       FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif'`
    ).all(),
    env2.DB.prepare(
      `SELECT id, activity_type, target_date FROM activity_schedule WHERE status = 'Pending' LIMIT 50`
    ).all(),
    env2.DB.prepare(
      `SELECT id, category, report_date FROM issues WHERE status != 'Done' LIMIT 50`
    ).all(),
    env2.DB.prepare(
      `SELECT id, submitter_name, submitted_at FROM supply_requests WHERE status = 'Pending' LIMIT 50`
    ).all()
  ]);
  const list4 = [];
  const today2 = /* @__PURE__ */ new Date();
  today2.setHours(0, 0, 0, 0);
  const intervals = [90, 60, 30, 14, 7, 1, 0];
  (contracts.results || []).forEach((c) => {
    if (!c.end_date)
      return;
    const end = /* @__PURE__ */ new Date(c.end_date + "T00:00:00");
    if (isNaN(end.getTime()))
      return;
    intervals.forEach((days) => {
      const remDate = new Date(end.getTime());
      remDate.setDate(end.getDate() - days);
      remDate.setHours(0, 0, 0, 0);
      const diffTime = today2.getTime() - remDate.getTime();
      const diffDays = diffTime / 864e5;
      if (diffDays >= 0 && diffDays <= 7) {
        const severity = days <= 7 ? "danger" : days <= 30 ? "warning" : "info";
        const title2 = days === 0 ? `Hari H: Kontrak ${c.emp_name} Berakhir!` : `Reminder H-${days}: Kontrak ${c.emp_name} Berakhir`;
        list4.push({
          id: `contract-${c.id}-${days}`,
          title: title2,
          type: "contract",
          date: remDate.toISOString().slice(0, 10),
          severity
        });
      }
    });
  });
  (pendingSched.results || []).forEach((s) => {
    list4.push({
      id: `schedule-${s.id}`,
      title: `Jadwal Pending: ${s.activity_type}`,
      type: "schedule",
      date: s.target_date,
      severity: "warning"
    });
  });
  (openIssues.results || []).forEach((i) => {
    list4.push({
      id: `issue-${i.id}`,
      title: `Permasalahan Baru: ${i.category}`,
      type: "issue",
      date: i.report_date,
      severity: "danger"
    });
  });
  (pendingSupply.results || []).forEach((s) => {
    list4.push({
      id: `supply-${s.id}`,
      title: `Permintaan Barang pending: ${s.submitter_name}`,
      type: "supply",
      date: String(s.submitted_at).slice(0, 10),
      severity: "info"
    });
  });
  list4.sort((a, b) => b.date.localeCompare(a.date));
  return ok(list4, 200, origin);
}
__name(getNotifications, "getNotifications");

// src/routes/import.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
init_auth();
async function handleImport(request, env2, origin) {
  const user = await authenticate(request, env2);
  if (!user)
    return unauthorized(origin);
  if (!hasPermission(user, "employees", "write"))
    return forbidden(origin);
  const url = new URL(request.url);
  const path = url.pathname.replace("/api/import", "");
  if (path === "/backup" && request.method === "GET") {
    return exportBackup(env2, origin);
  }
  if (path === "/sync-calendar" && request.method === "POST") {
    return triggerCalendarSync(env2, origin);
  }
  if (path === "/counts" && request.method === "GET") {
    const tables = [
      "validation_options",
      "employees",
      "contracts",
      "issues",
      "one_on_one",
      "training",
      "relievers",
      "activity_schedule",
      "inspection_reports",
      "cleaning_reports",
      "fogging_reports",
      "basecamp_reports",
      "sop",
      "master_checklist",
      "master_forms",
      "supply_requests"
    ];
    const counts = {};
    for (const t of tables) {
      try {
        const r = await env2.DB.prepare(`SELECT COUNT(*) as cnt FROM ${t}`).first();
        counts[t] = r.cnt;
      } catch (e) {
        counts[t] = -1;
      }
    }
    return ok(counts, 200, origin);
  }
  if (request.method !== "POST")
    return error3("Method not allowed", 405, origin);
  let body;
  try {
    body = await request.json();
  } catch {
    return error3("Invalid JSON", 400, origin);
  }
  const { rows = [], onDuplicate = "skip" } = body;
  if (!Array.isArray(rows))
    return error3("rows must be an array", 400, origin);
  try {
    switch (path) {
      case "/employees":
        return importEmployees2(rows, onDuplicate, env2, origin);
      case "/contracts":
        return importContracts2(rows, onDuplicate, env2, origin);
      case "/relievers":
        return importRelievers2(rows, onDuplicate, env2, origin);
      case "/schedule":
        return importSchedule2(rows, onDuplicate, env2, origin);
      case "/issues":
        return importIssues2(rows, onDuplicate, env2, origin);
      case "/one_on_one":
        return importOneOnOne(rows, onDuplicate, env2, origin);
      case "/training":
        return importTraining(rows, onDuplicate, env2, origin);
      case "/checklist":
        return importChecklist(rows, onDuplicate, env2, origin);
      case "/forms":
        return importForms(rows, onDuplicate, env2, origin);
      case "/sop":
        return importSop(rows, onDuplicate, env2, origin);
      case "/inspection":
        return importInspection(rows, onDuplicate, env2, origin);
      case "/cleaning":
        return importCleaning(rows, onDuplicate, env2, origin);
      case "/fogging":
        return importFogging(rows, onDuplicate, env2, origin);
      case "/basecamp":
        return importBasecamp(rows, onDuplicate, env2, origin);
      case "/supply":
        return importSupply(rows, onDuplicate, env2, origin);
      case "/validation":
        return importValidation(rows, onDuplicate, env2, origin);
      default:
        return error3("Unknown import module", 404, origin);
    }
  } catch (err) {
    return error3("Import error: " + err.message, 500, origin);
  }
}
__name(handleImport, "handleImport");
function safeStr(v) {
  return v !== void 0 && v !== null && String(v).trim() !== "" ? String(v).trim() : null;
}
__name(safeStr, "safeStr");
function safeDate(v) {
  return safeStr(v);
}
__name(safeDate, "safeDate");
function today() {
  return (/* @__PURE__ */ new Date()).toISOString().slice(0, 10);
}
__name(today, "today");
function makeBranchMatcher(branches) {
  return (str) => {
    if (!str)
      return null;
    const s = String(str).toLowerCase().trim();
    const exact = branches.find(
      (r) => r.full_name?.toLowerCase() === s || r.code?.toLowerCase() === s || r.name?.toLowerCase() === s
    );
    if (exact)
      return exact.id;
    const partial = branches.find(
      (r) => r.full_name?.toLowerCase().includes(s) || s.includes(r.code?.toLowerCase() || "~~~")
    );
    return partial ? partial.id : null;
  };
}
__name(makeBranchMatcher, "makeBranchMatcher");
async function batchInsert(db, stmts) {
  if (stmts.length === 0)
    return 0;
  const chunkSize = 100;
  let total = 0;
  for (let i = 0; i < stmts.length; i += chunkSize) {
    await db.batch(stmts.slice(i, i + chunkSize));
    total += stmts.slice(i, i + chunkSize).length;
  }
  return total;
}
__name(batchInsert, "batchInsert");
async function importValidation(rows, onDuplicate, env2, origin) {
  const branchSet = /* @__PURE__ */ new Set();
  const picSet = /* @__PURE__ */ new Set();
  const kegiatanSet = /* @__PURE__ */ new Set();
  const quartalSet = /* @__PURE__ */ new Set();
  const pkwtSet = /* @__PURE__ */ new Set();
  for (const row of rows) {
    if (row.cabang) {
      const val = String(row.cabang).trim();
      if (val)
        branchSet.add(val);
    }
    if (row.pic) {
      const val = String(row.pic).trim();
      if (val)
        picSet.add(val);
    }
    if (row.kegiatan) {
      const val = String(row.kegiatan).trim();
      if (val)
        kegiatanSet.add(val);
    }
    if (row.quartal) {
      const val = String(row.quartal).trim();
      if (val)
        quartalSet.add(val);
    }
    if (row.masa_pkwt) {
      const val = String(row.masa_pkwt).trim();
      if (val)
        pkwtSet.add(val);
    }
  }
  const db = env2.DB;
  const stmts = [];
  const activeBranchCodes = [];
  for (const full_name of branchSet) {
    let code = "";
    let name = "";
    const dotIdx = full_name.indexOf(".");
    if (dotIdx > 0) {
      code = full_name.slice(0, dotIdx).trim();
      name = full_name.slice(dotIdx + 1).trim();
    } else {
      code = full_name;
      name = full_name;
    }
    activeBranchCodes.push(code);
    stmts.push(db.prepare(
      `INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)
       ON CONFLICT(code) DO UPDATE SET name=excluded.name, full_name=excluded.full_name, is_active=1`
    ).bind(code, name, full_name));
  }
  if (activeBranchCodes.length > 0) {
    const placeholders = activeBranchCodes.map(() => "?").join(",");
    stmts.push(db.prepare(
      `UPDATE branches SET is_active = 0 WHERE code NOT IN (${placeholders})`
    ).bind(...activeBranchCodes));
  }
  const activePics = Array.from(picSet);
  for (const pic of activePics) {
    stmts.push(db.prepare(
      `INSERT INTO pic_list (name, is_active) VALUES (?, 1)
       ON CONFLICT(name) DO UPDATE SET is_active=1`
    ).bind(pic));
  }
  if (activePics.length > 0) {
    const placeholders = activePics.map(() => "?").join(",");
    stmts.push(db.prepare(
      `UPDATE pic_list SET is_active = 0 WHERE name NOT IN (${placeholders})`
    ).bind(...activePics));
  }
  stmts.push(db.prepare(`DELETE FROM validation_options`));
  for (const pic of activePics) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('pic', ?)`).bind(pic));
  }
  for (const keg of kegiatanSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('activity', ?)`).bind(keg));
  }
  for (const q of quartalSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('quarter', ?)`).bind(q));
  }
  for (const pkwt of pkwtSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('pkwt', ?)`).bind(pkwt));
  }
  await batchInsert(db, stmts);
  return ok({ inserted: branchSet.size + picSet.size + kegiatanSet.size + quartalSet.size + pkwtSet.size, skipped: 0 }, 200, origin);
}
__name(importValidation, "importValidation");
async function importEmployees2(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, full_name FROM employees").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((e) => {
    existingMap.set(e.full_name.toLowerCase().trim(), e.id);
  });
  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const full_name = safeStr(row.full_name);
    if (!full_name) {
      skipped++;
      continue;
    }
    importedNames.push(full_name);
    const key = full_name.toLowerCase().trim();
    const branch_id = matchBranch(row.branch_name);
    const division = safeStr(row.division) || "FACILITY CARE";
    const phone = safeStr(row.phone);
    const join_date = safeDate(row.join_date);
    const status = safeStr(row.status) || "Aktif";
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE employees SET branch_id = ?, division = ?, phone = ?, join_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, division, phone, join_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(full_name, branch_id, division, phone, join_date, status, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedNames.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, full_name FROM employees").all();
    const importedSet = new Set(importedNames.map((n) => n.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((e) => {
      if (!importedSet.has(e.full_name.toLowerCase().trim())) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM employees WHERE id = ?").bind(e.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importEmployees2, "importEmployees");
async function importContracts2(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const eRows = await env2.DB.prepare("SELECT id, full_name FROM employees").all();
  const matchEmployee = /* @__PURE__ */ __name((str) => {
    if (!str)
      return null;
    const s = str.toLowerCase().trim();
    const emp = eRows.results.find((r) => r.full_name.toLowerCase().trim() === s);
    return emp ? emp.id : null;
  }, "matchEmployee");
  const existing = await env2.DB.prepare("SELECT id, employee_name, start_date FROM contracts").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((c) => {
    if (c.employee_name && c.start_date) {
      existingMap.set(c.employee_name.toLowerCase().trim() + "_" + c.start_date, c.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const employee_name = safeStr(row.employee_name);
    if (!employee_name) {
      skipped++;
      continue;
    }
    const start_date = safeDate(row.start_date);
    const end_date = safeDate(row.end_date);
    const key = employee_name.toLowerCase().trim() + "_" + start_date;
    importedKeys.push(key);
    const employee_id = matchEmployee(employee_name);
    const branch_id = matchBranch(row.branch_name);
    const division = safeStr(row.division) || "FACILITY CARE";
    const contract_type = safeStr(row.contract_type);
    const pkwt_number = safeStr(row.pkwt_number);
    const status = safeStr(row.status) || "Aktif";
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE contracts SET employee_id = ?, branch_id = ?, division = ?, end_date = ?, contract_type = ?, pkwt_number = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(employee_id, branch_id, division, end_date, contract_type, pkwt_number, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO contracts (employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, employee_name, start_date FROM contracts").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((c) => {
      const key = (c.employee_name || "").toLowerCase().trim() + "_" + c.start_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM contracts WHERE id = ?").bind(c.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importContracts2, "importContracts");
async function importRelievers2(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, reliever_name, backup_date FROM relievers").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((r) => {
    if (r.reliever_name && r.backup_date) {
      existingMap.set(r.reliever_name.toLowerCase().trim() + "_" + r.backup_date, r.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const reliever_name = safeStr(row.reliever_name) || "-";
    const backup_date = safeDate(row.backup_date) || today();
    const key = reliever_name.toLowerCase().trim() + "_" + backup_date;
    importedKeys.push(key);
    const branch_id = matchBranch(row.branch_name);
    const original_fc_name = safeStr(row.original_fc_name);
    const period = safeStr(row.period);
    const completion_date = safeDate(row.completion_date);
    const reason = safeStr(row.reason);
    const shift = safeStr(row.shift);
    const status = safeStr(row.status) || "Pending";
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE relievers SET branch_id = ?, original_fc_name = ?, period = ?, completion_date = ?, reason = ?, shift = ?, status = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, original_fc_name, period, completion_date, reason, shift, status, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, reliever_name, backup_date FROM relievers").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((r) => {
      const key = (r.reliever_name || "").toLowerCase().trim() + "_" + r.backup_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM relievers WHERE id = ?").bind(r.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importRelievers2, "importRelievers");
async function importSchedule2(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, activity_type, target_date, branch_id FROM activity_schedule").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((s) => {
    if (s.activity_type && s.target_date) {
      existingMap.set(s.activity_type.toLowerCase().trim() + "_" + s.target_date + "_" + s.branch_id, s.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const activity_type = safeStr(row.activity_type);
    if (!activity_type) {
      skipped++;
      continue;
    }
    const target_date = safeDate(row.target_date);
    const branch_id = matchBranch(row.branch_name);
    const key = activity_type.toLowerCase().trim() + "_" + target_date + "_" + branch_id;
    importedKeys.push(key);
    const period = safeStr(row.period);
    const pic = safeStr(row.pic);
    const opening_date = safeDate(row.opening_date);
    const completion_date = safeDate(row.completion_date);
    const status = safeStr(row.status) || "Pending";
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE activity_schedule SET period = ?, pic = ?, opening_date = ?, completion_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(period, pic, opening_date, completion_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, activity_type, target_date, branch_id FROM activity_schedule").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((s) => {
      const key = (s.activity_type || "").toLowerCase().trim() + "_" + s.target_date + "_" + s.branch_id;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM activity_schedule WHERE id = ?").bind(s.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importSchedule2, "importSchedule");
async function importIssues2(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, complaint, report_date, branch_id FROM issues").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((i) => {
    if (i.complaint && i.report_date) {
      existingMap.set(i.complaint.toLowerCase().trim() + "_" + i.report_date + "_" + i.branch_id, i.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const complaint = safeStr(row.complaint);
    if (!complaint) {
      skipped++;
      continue;
    }
    const report_date = safeDate(row.report_date) || today();
    const branch_id = matchBranch(row.branch_name);
    const key = complaint.toLowerCase().trim() + "_" + report_date + "_" + branch_id;
    importedKeys.push(key);
    const category = safeStr(row.category) || "Lainnya";
    const source = safeStr(row.source);
    const employee_name = safeStr(row.employee_name);
    const fc_specialist = safeStr(row.fc_specialist);
    const solution = safeStr(row.solution);
    const status = safeStr(row.status) || "Open";
    const completion_date = safeDate(row.completion_date);
    let day_count = null;
    if (completion_date && report_date) {
      const d1 = new Date(report_date), d2 = new Date(completion_date);
      if (!isNaN(d1) && !isNaN(d2))
        day_count = Math.floor((d2 - d1) / 864e5);
    }
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE issues SET category = ?, source = ?, employee_name = ?, fc_specialist = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, source, employee_name, fc_specialist, solution, status, completion_date, day_count, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, complaint, report_date, branch_id FROM issues").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((i) => {
      const key = (i.complaint || "").toLowerCase().trim() + "_" + i.report_date + "_" + i.branch_id;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM issues WHERE id = ?").bind(i.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importIssues2, "importIssues");
async function importOneOnOne(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, employee_name, meeting_date FROM one_on_one").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((o) => {
    if (o.employee_name && o.meeting_date) {
      existingMap.set(o.employee_name.toLowerCase().trim() + "_" + o.meeting_date, o.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const employee_name = safeStr(row.employee_name) || "-";
    const meeting_date = safeDate(row.meeting_date) || today();
    const key = employee_name.toLowerCase().trim() + "_" + meeting_date;
    importedKeys.push(key);
    const branch_id = matchBranch(row.branch_name);
    const pic = safeStr(row.pic);
    const problem = safeStr(row.problem) || "-";
    const solution = safeStr(row.solution);
    const status = safeStr(row.status) || "Open";
    const completion_date = safeDate(row.completion_date);
    const document_link = safeStr(row.document_link);
    let day_count = null;
    if (completion_date && meeting_date) {
      const d1 = new Date(meeting_date), d2 = new Date(completion_date);
      if (!isNaN(d1) && !isNaN(d2))
        day_count = Math.floor((d2 - d1) / 864e5);
    }
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE one_on_one SET branch_id = ?, pic = ?, problem = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, document_link = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, pic, problem, solution, status, completion_date, day_count, document_link, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, employee_name, meeting_date FROM one_on_one").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((o) => {
      const key = (o.employee_name || "").toLowerCase().trim() + "_" + o.meeting_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM one_on_one WHERE id = ?").bind(o.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importOneOnOne, "importOneOnOne");
async function importTraining(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, subject, training_date, batch FROM training").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((t) => {
    if (t.subject && t.training_date) {
      existingMap.set(t.subject.toLowerCase().trim() + "_" + t.training_date + "_" + (t.batch || "").toLowerCase().trim(), t.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const subject = safeStr(row.subject);
    if (!subject) {
      skipped++;
      continue;
    }
    const training_date = safeDate(row.training_date) || today();
    const batch = safeStr(row.batch) || "";
    const key = subject.toLowerCase().trim() + "_" + training_date + "_" + batch.toLowerCase().trim();
    importedKeys.push(key);
    const branch_id = matchBranch(row.branch_name);
    const participants = safeStr(row.participants);
    const trainer = safeStr(row.trainer);
    const score = row.score != null ? parseFloat(row.score) : null;
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE training SET branch_id = ?, participants = ?, trainer = ?, score = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, participants, trainer, score, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(training_date, batch, subject, participants, branch_id, trainer, score, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, subject, training_date, batch FROM training").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((t) => {
      const key = (t.subject || "").toLowerCase().trim() + "_" + t.training_date + "_" + (t.batch || "").toLowerCase().trim();
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM training WHERE id = ?").bind(t.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importTraining, "importTraining");
async function importChecklist(rows, onDuplicate, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id, name FROM master_checklist").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((c) => {
    existingMap.set(c.name.toLowerCase().trim(), c.id);
  });
  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name) || "-";
    importedNames.push(name);
    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || "Umum";
    const document_link = safeStr(row.document_link);
    const description = safeStr(row.description);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE master_checklist SET category = ?, document_link = ?, description = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, description, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO master_checklist (name, category, document_link, description) VALUES (?, ?, ?, ?)`
      ).bind(name, category, document_link, description));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedNames.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, name FROM master_checklist").all();
    const importedSet = new Set(importedNames.map((n) => n.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((c) => {
      if (!importedSet.has(c.name.toLowerCase().trim())) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM master_checklist WHERE id = ?").bind(c.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importChecklist, "importChecklist");
async function importForms(rows, onDuplicate, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id, name FROM master_forms").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((f) => {
    existingMap.set(f.name.toLowerCase().trim(), f.id);
  });
  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) {
      skipped++;
      continue;
    }
    importedNames.push(name);
    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || "Umum";
    const document_link = safeStr(row.document_link);
    const description = safeStr(row.description);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE master_forms SET category = ?, document_link = ?, description = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, description, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO master_forms (name, category, document_link, description) VALUES (?, ?, ?, ?)`
      ).bind(name, category, document_link, description));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedNames.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, name FROM master_forms").all();
    const importedSet = new Set(importedNames.map((n) => n.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((f) => {
      if (!importedSet.has(f.name.toLowerCase().trim())) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM master_forms WHERE id = ?").bind(f.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importForms, "importForms");
async function importSop(rows, onDuplicate, env2, origin) {
  const existing = await env2.DB.prepare("SELECT id, name FROM sop").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((s) => {
    existingMap.set(s.name.toLowerCase().trim(), s.id);
  });
  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) {
      skipped++;
      continue;
    }
    importedNames.push(name);
    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || "Umum";
    const document_link = safeStr(row.document_link);
    const version2 = safeStr(row.version) || "1.0";
    const effective_date = safeDate(row.effective_date);
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE sop SET category = ?, document_link = ?, version = ?, effective_date = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, version2, effective_date, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO sop (name, category, document_link, version, effective_date, notes) VALUES (?, ?, ?, ?, ?, ?)`
      ).bind(name, category, document_link, version2, effective_date, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedNames.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, name FROM sop").all();
    const importedSet = new Set(importedNames.map((n) => n.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((s) => {
      if (!importedSet.has(s.name.toLowerCase().trim())) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM sop WHERE id = ?").bind(s.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importSop, "importSop");
async function importInspection(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, branch_id, period, inspection_date FROM inspection_reports").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((i) => {
    if (i.inspection_date) {
      existingMap.set(i.branch_id + "_" + (i.period || "").toLowerCase().trim() + "_" + i.inspection_date, i.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) {
      skipped++;
      continue;
    }
    const branch_id = matchBranch(branch_name);
    const period = safeStr(row.period) || "-";
    const inspection_date = safeDate(row.inspection_date) || today();
    const key = branch_id + "_" + period.toLowerCase().trim() + "_" + inspection_date;
    importedKeys.push(key);
    const status = safeStr(row.status) || "Pending";
    const fc_score = row.fc_score != null ? parseFloat(row.fc_score) : null;
    const spv_score = row.spv_score != null ? parseFloat(row.spv_score) : null;
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE inspection_reports SET status = ?, fc_score = ?, spv_score = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, fc_score, spv_score, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, branch_id, period, inspection_date FROM inspection_reports").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((i) => {
      const key = i.branch_id + "_" + (i.period || "").toLowerCase().trim() + "_" + i.inspection_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM inspection_reports WHERE id = ?").bind(i.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importInspection, "importInspection");
async function importCleaning(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, branch_id, activity_type, period, activity_date FROM cleaning_reports").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((c) => {
    if (c.activity_date) {
      existingMap.set(c.branch_id + "_" + c.activity_type.toLowerCase().trim() + "_" + (c.period || "").toLowerCase().trim() + "_" + c.activity_date, c.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) {
      skipped++;
      continue;
    }
    const branch_id = matchBranch(branch_name);
    const activity_type = safeStr(row.activity_type) || "General Cleaning";
    const period = safeStr(row.period) || "-";
    const activity_date = safeDate(row.activity_date) || today();
    const key = branch_id + "_" + activity_type.toLowerCase().trim() + "_" + period.toLowerCase().trim() + "_" + activity_date;
    importedKeys.push(key);
    const status = safeStr(row.status) || "Pending";
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE cleaning_reports SET status = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, activity_type, period, activity_date, status, document_link, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, branch_id, activity_type, period, activity_date FROM cleaning_reports").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((c) => {
      const key = c.branch_id + "_" + c.activity_type.toLowerCase().trim() + "_" + (c.period || "").toLowerCase().trim() + "_" + c.activity_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM cleaning_reports WHERE id = ?").bind(c.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importCleaning, "importCleaning");
async function importFogging(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, branch_id, period, activity_date FROM fogging_reports").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((f) => {
    if (f.activity_date) {
      existingMap.set(f.branch_id + "_" + (f.period || "").toLowerCase().trim() + "_" + f.activity_date, f.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) {
      skipped++;
      continue;
    }
    const branch_id = matchBranch(branch_name);
    const period = safeStr(row.period) || "-";
    const activity_date = safeDate(row.activity_date) || today();
    const key = branch_id + "_" + period.toLowerCase().trim() + "_" + activity_date;
    importedKeys.push(key);
    const status = safeStr(row.status) || "Pending";
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE fogging_reports SET status = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, "Fogging", period, activity_date, status, document_link, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, branch_id, period, activity_date FROM fogging_reports").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((f) => {
      const key = f.branch_id + "_" + (f.period || "").toLowerCase().trim() + "_" + f.activity_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM fogging_reports WHERE id = ?").bind(f.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importFogging, "importFogging");
async function importBasecamp(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, branch_id, problem, info_date FROM basecamp_reports").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((b) => {
    if (b.info_date) {
      existingMap.set(b.branch_id + "_" + b.problem.toLowerCase().trim() + "_" + b.info_date, b.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const problem = safeStr(row.problem);
    if (!problem) {
      skipped++;
      continue;
    }
    const branch_id = matchBranch(row.branch_name);
    const info_date = safeDate(row.info_date) || today();
    const key = branch_id + "_" + problem.toLowerCase().trim() + "_" + info_date;
    importedKeys.push(key);
    const pic = safeStr(row.pic);
    const done_date = safeDate(row.done_date);
    const status = safeStr(row.status) || "Open";
    const notes = safeStr(row.notes);
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE basecamp_reports SET pic = ?, done_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(pic, done_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, problem, pic, info_date, done_date, status, notes));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, branch_id, problem, info_date FROM basecamp_reports").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((b) => {
      const key = b.branch_id + "_" + b.problem.toLowerCase().trim() + "_" + b.info_date;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM basecamp_reports WHERE id = ?").bind(b.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importBasecamp, "importBasecamp");
async function importSupply(rows, onDuplicate, env2, origin) {
  const bRows = await env2.DB.prepare("SELECT id, code, name, full_name FROM branches WHERE is_active = 1").all();
  const matchBranch = makeBranchMatcher(bRows.results);
  const existing = await env2.DB.prepare("SELECT id, submitter_name, submitted_at FROM supply_requests").all();
  const existingMap = /* @__PURE__ */ new Map();
  (existing.results || []).forEach((s) => {
    if (s.submitter_name && s.submitted_at) {
      existingMap.set(s.submitter_name.toLowerCase().trim() + "_" + s.submitted_at, s.id);
    }
  });
  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  for (const row of rows) {
    const submitter_name = safeStr(row.submitter_name) || "-";
    const submitted_at = safeDate(row.submitted_at) || today();
    const key = submitter_name.toLowerCase().trim() + "_" + submitted_at;
    importedKeys.push(key);
    const branch_id = matchBranch(row.branch_name);
    const branchObj = branch_id ? bRows.results.find((b) => b.id === branch_id) : null;
    const branch_name = branchObj ? branchObj.full_name : safeStr(row.branch_name);
    const tools_items = safeStr(row.tools_items);
    const tools_quantity = safeStr(row.tools_quantity);
    const chemical_items = safeStr(row.chemical_items);
    const chemical_quantity = safeStr(row.chemical_quantity);
    const additional_notes = safeStr(row.additional_notes);
    const status = safeStr(row.status) || "Pending";
    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === "update") {
        stmts.push(env2.DB.prepare(
          `UPDATE supply_requests SET branch_id = ?, branch_name = ?, tools_items = ?, tools_quantity = ?, chemical_items = ?, chemical_quantity = ?, additional_notes = ?, status = ?, processed_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env2.DB.prepare(
        `INSERT INTO supply_requests (submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status));
      inserted++;
    }
  }
  await batchInsert(env2.DB, stmts);
  if (importedKeys.length > 0) {
    const currentAll = await env2.DB.prepare("SELECT id, submitter_name, submitted_at FROM supply_requests").all();
    const importedSet = new Set(importedKeys.map((k) => k.toLowerCase().trim()));
    const deleteStmts = [];
    (currentAll.results || []).forEach((s) => {
      const key = (s.submitter_name || "").toLowerCase().trim() + "_" + s.submitted_at;
      if (!importedSet.has(key)) {
        deleteStmts.push(env2.DB.prepare("DELETE FROM supply_requests WHERE id = ?").bind(s.id));
      }
    });
    if (deleteStmts.length > 0)
      await batchInsert(env2.DB, deleteStmts);
  }
  return ok({ inserted, skipped, updated }, 200, origin);
}
__name(importSupply, "importSupply");
async function exportBackup(env2, origin) {
  const tables = [
    "users",
    "branches",
    "employees",
    "contracts",
    "activity_schedule",
    "issues",
    "one_on_one",
    "training",
    "relievers",
    "inspection_reports",
    "cleaning_reports",
    "fogging_reports",
    "basecamp_reports",
    "supply_requests",
    "master_checklist",
    "master_forms",
    "sop"
  ];
  const promises = tables.map(async (table3) => {
    try {
      const res = await env2.DB.prepare(`SELECT * FROM ${table3}`).all();
      return { table: table3, rows: res.results || [] };
    } catch (e) {
      return { table: table3, error: e.message };
    }
  });
  const results = await Promise.all(promises);
  const backup = {};
  results.forEach((item) => {
    backup[item.table] = item.rows;
  });
  return ok({
    timestamp: (/* @__PURE__ */ new Date()).toISOString(),
    database: backup
  }, 200, origin);
}
__name(exportBackup, "exportBackup");
async function triggerCalendarSync(env2, origin) {
  const { getSyncStatements: getSyncStatements2 } = await Promise.resolve().then(() => (init_calendar(), calendar_exports));
  const allStmts = [];
  allStmts.push(env2.DB.prepare("DELETE FROM calendar_events"));
  const contr = await env2.DB.prepare(`SELECT c.id, c.end_date, c.status, c.branch_id, COALESCE(e.full_name, c.employee_name, '?') as empName FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id`).all();
  (contr.results || []).forEach((c) => {
    allStmts.push(...getSyncStatements2(env2.DB, "contracts", c.id, c));
  });
  const sched = await env2.DB.prepare(`SELECT * FROM activity_schedule`).all();
  (sched.results || []).forEach((s) => {
    allStmts.push(...getSyncStatements2(env2.DB, "schedule", s.id, s));
  });
  const rel = await env2.DB.prepare(`SELECT * FROM relievers`).all();
  (rel.results || []).forEach((r) => {
    allStmts.push(...getSyncStatements2(env2.DB, "relievers", r.id, r));
  });
  const iss = await env2.DB.prepare(`SELECT * FROM issues`).all();
  (iss.results || []).forEach((i) => {
    allStmts.push(...getSyncStatements2(env2.DB, "issues", i.id, i));
  });
  const ooo = await env2.DB.prepare(`SELECT * FROM one_on_one`).all();
  (ooo.results || []).forEach((o) => {
    allStmts.push(...getSyncStatements2(env2.DB, "one_on_one", o.id, o));
  });
  const tr = await env2.DB.prepare(`SELECT * FROM training`).all();
  (tr.results || []).forEach((t) => {
    allStmts.push(...getSyncStatements2(env2.DB, "training", t.id, t));
  });
  const clean = await env2.DB.prepare(`SELECT * FROM cleaning_reports`).all();
  (clean.results || []).forEach((c) => {
    allStmts.push(...getSyncStatements2(env2.DB, "cleaning", c.id, c));
  });
  const insp = await env2.DB.prepare(`SELECT * FROM inspection_reports`).all();
  (insp.results || []).forEach((i) => {
    allStmts.push(...getSyncStatements2(env2.DB, "inspection", i.id, i));
  });
  const fog = await env2.DB.prepare(`SELECT * FROM fogging_reports`).all();
  (fog.results || []).forEach((f) => {
    allStmts.push(...getSyncStatements2(env2.DB, "fogging", f.id, f));
  });
  const base = await env2.DB.prepare(`SELECT * FROM basecamp_reports`).all();
  (base.results || []).forEach((b) => {
    allStmts.push(...getSyncStatements2(env2.DB, "basecamp", b.id, b));
  });
  const supply = await env2.DB.prepare(`SELECT * FROM supply_requests`).all();
  (supply.results || []).forEach((s) => {
    allStmts.push(...getSyncStatements2(env2.DB, "supply", s.id, s));
  });
  if (allStmts.length > 0) {
    const chunkSize = 100;
    for (let i = 0; i < allStmts.length; i += chunkSize) {
      await env2.DB.batch(allStmts.slice(i, i + chunkSize));
    }
  }
  return ok({ message: `Berhasil sinkronisasi ${allStmts.length} event kalender dari data lama.` }, 200, origin);
}
__name(triggerCalendarSync, "triggerCalendarSync");

// src/utils/google_sync.js
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_process();
init_virtual_unenv_global_polyfill_cloudflare_unenv_preset_node_console();
init_performance2();
var import_papaparse = __toESM(require_papaparse_min());
var SPREADSHEET_ID = "1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o";
async function fetchCSV(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const res = await fetch(url);
  if (!res.ok)
    throw new Error(`Failed to fetch sheet ${sheetName}`);
  const text = await res.text();
  const parsed = import_papaparse.default.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}
__name(fetchCSV, "fetchCSV");
async function syncGoogleSheets(env2) {
  try {
    console.log("Starting Google Sheets Sync...");
    const empData = await fetchCSV("Master Karyawan");
    const valData = await fetchCSV("Validasi");
    const branches = /* @__PURE__ */ new Set();
    empData.forEach((r) => {
      if (r["Cabang"])
        branches.add(r["Cabang"].trim());
    });
    valData.forEach((r) => {
      if (r["CABANG"])
        branches.add(r["CABANG"].trim());
    });
    const currentBranches = (await env2.DB.prepare("SELECT id, full_name FROM branches").all()).results;
    for (const b of branches) {
      if (!b)
        continue;
      const existing = currentBranches.find((cb) => cb.full_name === b);
      if (!existing) {
        await env2.DB.prepare("INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)").bind(b.split(".")[0] || "", b, b).run();
      }
    }
    const updatedBranches = (await env2.DB.prepare("SELECT id, full_name FROM branches").all()).results;
    const branchMap = {};
    updatedBranches.forEach((b) => branchMap[b.full_name] = b.id);
    const empStmts = [];
    const currentEmp = (await env2.DB.prepare("SELECT id, full_name FROM employees").all()).results;
    const sheetEmpNames = /* @__PURE__ */ new Set();
    for (const row of empData) {
      const name = (row["Nama Lengkap"] || "").trim();
      if (!name)
        continue;
      sheetEmpNames.add(name);
      const bName = (row["Cabang"] || "").trim();
      const bId = branchMap[bName] || null;
      const existing = currentEmp.find((e) => e.full_name === name);
      if (existing) {
        empStmts.push(env2.DB.prepare('UPDATE employees SET branch_id=?, status=?, updated_at=datetime("now") WHERE id=?').bind(bId, "Aktif", existing.id));
      } else {
        empStmts.push(env2.DB.prepare("INSERT INTO employees (full_name, branch_id, status) VALUES (?, ?, ?)").bind(name, bId, "Aktif"));
      }
    }
    valData.forEach((row) => {
      const name = (row["NAMA KARYAWAN"] || "").trim();
      if (name && !sheetEmpNames.has(name)) {
        sheetEmpNames.add(name);
        const existing = currentEmp.find((e) => e.full_name === name);
        if (!existing) {
          empStmts.push(env2.DB.prepare("INSERT INTO employees (full_name, branch_id, status) VALUES (?, ?, ?)").bind(name, branchMap[(row["CABANG"] || "").trim()] || null, "Aktif"));
        } else {
          empStmts.push(env2.DB.prepare('UPDATE employees SET status=?, updated_at=datetime("now") WHERE id=?').bind("Aktif", existing.id));
        }
      }
    });
    for (let i = 0; i < empStmts.length; i += 100) {
      await env2.DB.batch(empStmts.slice(i, i + 100));
    }
    const picStmts = [];
    const currentPics = (await env2.DB.prepare("SELECT id, name, role FROM pic_list").all()).results;
    for (const row of valData) {
      const pic = (row["PIC"] || "").trim();
      const role = (row["KEGIATAN"] || "").trim();
      if (pic) {
        const existing = currentPics.find((p) => p.name === pic);
        if (!existing) {
          picStmts.push(env2.DB.prepare("INSERT INTO pic_list (name, role) VALUES (?, ?)").bind(pic, role));
          currentPics.push({ name: pic, role });
        }
      }
      const picPelapor = (row["PIC PELAPOR"] || "").trim();
      if (picPelapor) {
        const existingPelapor = currentPics.find((p) => p.name === picPelapor);
        if (!existingPelapor) {
          picStmts.push(env2.DB.prepare("INSERT INTO pic_list (name, role) VALUES (?, ?)").bind(picPelapor, "Pelapor"));
          currentPics.push({ name: picPelapor, role: "Pelapor" });
        }
      }
    }
    for (let i = 0; i < picStmts.length; i += 100) {
      await env2.DB.batch(picStmts.slice(i, i + 100));
    }
    console.log("Google Sheets Sync Complete!");
    return { success: true, message: `Berhasil sinkronisasi Google Sheets. (Ditambahkan/Diupdate: ${empStmts.length} Karyawan, ${picStmts.length} PIC). Data FCMS lama tetap aman.` };
  } catch (err) {
    console.error("Google Sheets Sync Error:", err);
    return { success: false, error: err.message };
  }
}
__name(syncGoogleSheets, "syncGoogleSheets");

// src/index.js
var src_default = {
  async fetch(request, env2, ctx) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || env2.CORS_ORIGIN || "*";
    if (request.method === "OPTIONS") {
      return options(origin);
    }
    if (!url.pathname.startsWith("/api/")) {
      return new Response("FM Operations API", {
        status: 200,
        headers: { "Content-Type": "text/plain" }
      });
    }
    try {
      const path = url.pathname;
      if (path.startsWith("/api/auth"))
        return handleAuth(request, env2, origin);
      if (path.startsWith("/api/users"))
        return handleUsers(request, env2, origin);
      if (path.startsWith("/api/branches"))
        return handleBranches(request, env2, origin);
      if (path.startsWith("/api/employees"))
        return handleEmployees(request, env2, origin);
      if (path.startsWith("/api/contracts"))
        return handleContracts(request, env2, origin);
      if (path.startsWith("/api/schedule"))
        return handleSchedule(request, env2, origin);
      if (path.startsWith("/api/issues"))
        return handleIssues(request, env2, origin);
      if (path.startsWith("/api/one-on-one"))
        return handleOneOnOne(request, env2, origin);
      if (path.startsWith("/api/training"))
        return handleTraining(request, env2, origin);
      if (path.startsWith("/api/relievers"))
        return handleRelievers(request, env2, origin);
      if (path.startsWith("/api/reports"))
        return handleReports(request, env2, origin);
      if (path.startsWith("/api/dashboard"))
        return handleDashboard(request, env2, origin);
      if (path.startsWith("/api/import"))
        return handleImport(request, env2, origin);
      if (path === "/api/sync/google-sheets" && request.method === "POST") {
        const result = await syncGoogleSheets(env2);
        return new Response(JSON.stringify(result), {
          status: result.success ? 200 : 500,
          headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": origin }
        });
      }
      if (path.startsWith("/api/sop") || path.startsWith("/api/checklist") || path.startsWith("/api/forms") || path.startsWith("/api/pic")) {
        return handleMisc(request, env2, origin);
      }
      return error3("API endpoint not found", 404, origin);
    } catch (err) {
      console.error("Worker error:", err);
      return error3("Internal server error: " + err.message, 500, origin);
    }
  },
  async scheduled(event, env2, ctx) {
    ctx.waitUntil(syncGoogleSheets(env2));
  }
};
export {
  src_default as default
};
/*! Bundled license information:

papaparse/papaparse.min.js:
  (* @license
  Papa Parse
  v5.5.4
  https://github.com/mholt/PapaParse
  License: MIT
  *)
*/
//# sourceMappingURL=index.js.map
