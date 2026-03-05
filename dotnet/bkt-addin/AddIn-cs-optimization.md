# BKT COM Add-in (.NET) Optimization Recommendations

This document outlines optimization opportunities for the C# part of the BKT COM Add-in to improve startup performance and runtime efficiency.

---

## 🚀 HIGH IMPACT - Configuration/Build Optimizations

### 1. NGEN Pre-compilation (Native Image Generation)

**Problem:** IronPython and the DLR (Dynamic Language Runtime) assemblies are JIT-compiled on every Office startup.

**Solution:** Pre-compile assemblies to native images during installation using NGEN.

**Impact:** Can reduce .NET startup time by 30-50%.

**Implementation:** Add NGEN commands to installer scripts:

```batch
@echo off
set NGEN=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\ngen.exe

echo Installing native images...
%NGEN% install "%~dp0..\bin\BKT.dll"
%NGEN% install "%~dp0..\bin\IronPython.dll"
%NGEN% install "%~dp0..\bin\IronPython.Modules.dll"
%NGEN% install "%~dp0..\bin\Microsoft.Scripting.dll"
%NGEN% install "%~dp0..\bin\Microsoft.Dynamic.dll"
echo Done.
```

For uninstall, add:
```batch
%NGEN% uninstall BKT
%NGEN% uninstall IronPython
%NGEN% uninstall IronPython.Modules
%NGEN% uninstall Microsoft.Scripting
%NGEN% uninstall Microsoft.Dynamic
```

---

### 2. Assembly Binding Redirects Update

**Problem:** The `app.config` has binding redirects pointing to old versions (`1.3.3.0`) while using DLR `1.3.5.0`.

**Current (incorrect):**
```xml
<bindingRedirect oldVersion="0.0.0.0-1.3.3.0" newVersion="1.3.3.0" />
```

**Should be:**
```xml
<bindingRedirect oldVersion="0.0.0.0-1.3.5.0" newVersion="1.3.5.0" />
```

**Impact:** Eliminates assembly resolution overhead and potential runtime binding failures.

**Files to update:** `dotnet/bkt-addin/app.config`

---

### 3. Release Build Configuration Enhancements

**Problem:** Current Release config enables `Optimize=True` but could be more aggressive.

**Potential additions to `app.config`:**

```xml
<configuration>
  <runtime>
    <gcServer enabled="true"/>
    <gcConcurrent enabled="true"/>
    <!-- Reduces GC pauses in multi-threaded scenarios -->
  </runtime>
</configuration>
```

**Note:** Test thoroughly as `gcServer` may increase memory usage.

---

## ⚡ MEDIUM IMPACT - Startup Flow Optimizations

### 4. Async Startup Enhancement

**Current behavior:** 
- `async_startup=True` runs `BootstrapAddIn()` on a background thread
- However, `CreatePythonEngine()` is still synchronous in `OnConnection2`

**Opportunity:** Move IronPython engine creation itself to async, not just the bootstrap phase.

**Concept:**
```csharp
public void OnConnection2(object application)
{
    // Quick synchronous setup
    DetermineHostApplication(application);
    ReloadConfig();
    Reset();
    this.app = application;
    
    if (async_startup) {
        // Move EVERYTHING to background
        Thread bootstrapperThread = new Thread(() => {
            LoadPython();      // <-- Move this inside
            AsyncStartup();
        });
        bootstrapperThread.Start();
    } else {
        LoadPython();
        BootstrapAddIn();
        created = true;
    }
}
```

**Benefit:** Office ribbon appears faster; Python initializes completely in background.

---

### 5. Lazy Hook Subscription

**Problem:** Mouse/keyboard hooks (`HookEvents()`) are subscribed in the constructor before Office is fully loaded.

**Current code location:** `AddIn()` constructor, lines ~115-130

**Current (hooks in constructor):**
```csharp
public AddIn() {
    // ... logging setup ...
    
    // Initialize Mouse/Key-Hooks
    try {
        bool use_keymouse_hooks = Boolean.Parse(GetConfigEntry("use_keymouse_hooks", "true"));
        if (use_keymouse_hooks) {
            HookEvents();  // <-- Called too early
        }
    }
}
```

**Recommended:** Defer to `OnConnection2` or after `created=true`:
```csharp
public AddIn() {
    // ... logging setup only ...
    // Remove HookEvents() call from here
}

// Later, after Python is ready:
private void AsyncStartup() {
    // ... existing code ...
    created = true;
    
    // Now subscribe to hooks
    if (use_keymouse_hooks) {
        HookEvents();
    }
}
```

**Benefit:** Faster constructor execution, reduced memory footprint initially.

---

### 6. Conditional Debug Logging Optimization

**Problem:** Even in Release builds, string concatenation happens before `DebugMessage()` is called.

**Current pattern:**
```csharp
DebugMessage("event GetEnabled " + control.Id);  // String concat always happens
```

**The `[Conditional("DEBUG")]` attribute** eliminates the method call in Release, but the string is still built.

**Better pattern:**
```csharp
[Conditional("DEBUG")]
private void DebugMessage(string s) {
    Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + s);
}

// For hot paths, use interpolation (compiler can optimize better):
DebugMessage($"event GetEnabled {control.Id}");

// Or for truly zero-cost in Release, check first:
#if DEBUG
    DebugMessage("event GetEnabled " + control.Id);
#endif
```

**Impact:** Minor but adds up across hundreds of callback invocations.

---

## 🔧 MEDIUM IMPACT - Structural Improvements

### 7. Python Delegate Caching / Call Batching

**Problem:** Each ribbon callback makes a separate IronPython call.

**Current flow per control:**
```
GetEnabled → IronPython → return
GetLabel → IronPython → return
GetImage → IronPython → return
GetVisible → IronPython → return
... (repeat for each control)
```

**Opportunity:** Implement callback result caching on C# side:

```csharp
private Dictionary<string, object> callbackCache = new Dictionary<string, object>();
private DateTime lastCacheInvalidation = DateTime.MinValue;

public bool PythonGetEnabled(IRibbonControl control)
{
    string cacheKey = $"enabled_{control.Id}";
    if (callbackCache.TryGetValue(cacheKey, out object cached))
    {
        return (bool)cached;
    }
    
    bool result = python_delegate.get_enabled(control) == true;
    callbackCache[cacheKey] = result;
    return result;
}

// Clear cache on ribbon.Invalidate()
public void InvalidateCache()
{
    callbackCache.Clear();
}
```

**Note:** Requires coordination with Python side for proper invalidation.

---

### 8. Streamline Host Application Detection

**Problem:** `DetermineHostApplication()` uses sequential `is` checks with try-catch.

**Current:**
```csharp
if (application is Excel.Application) { ... }
else if (application is PowerPoint.Application) { ... }
else if (application is Word.Application) { ... }
// etc.
```

**Alternative using Type name (may be faster):**
```csharp
string typeName = application.GetType().FullName;
if (typeName.Contains("Excel")) {
    host = HostApplication.Excel;
    // ...
}
```

**Impact:** Minor, but cleaner and avoids potential COM interop overhead from `is` checks.

---

## 📦 BUILD/DEPLOYMENT Optimizations

### 9. Strong Name Signing

**Problem:** `SignAssembly=True` in csproj but key file reference is commented out.

**Current in bkt-addin.csproj:**
```xml
<SignAssembly>True</SignAssembly>
<!-- <AssemblyOriginatorKeyFile>bkt-addin-debug.snk</AssemblyOriginatorKeyFile> -->
```

**Impact:** Unsigned assemblies cannot be NGEN'd to the GAC for maximum performance.

**Solution:** Uncomment and ensure the .snk file exists:
```xml
<SignAssembly>True</SignAssembly>
<AssemblyOriginatorKeyFile>bkt-addin-debug.snk</AssemblyOriginatorKeyFile>
```

---

### 10. Conditional Package Loading

**Problem:** MahApps.Metro and Fluent.Ribbon are loaded even when TaskPanes are disabled.

**Current:** Fluent.dll is loaded in `CTPFactoryAvailable()`:
```csharp
string path = Path.Combine(codebase, "Fluent.dll");
Assembly.LoadFrom(path);
```

**This is already conditional** (only loads if `task_panes=true`), which is good.

**Verify:** Ensure other WPF-related assemblies aren't loaded unnecessarily at startup.

---

## 📊 Recommended Priority Order

| Priority | Optimization | Effort | Expected Impact |
|----------|-------------|--------|-----------------|
| 1 | NGEN Pre-compilation | Low | High (30-50% startup reduction) |
| 2 | Fix Assembly Binding Redirects | Low | Medium (eliminates binding failures) |
| 3 | Move IronPython creation to async | Medium | High (perceived startup speed) |
| 4 | Lazy hook subscription | Low | Low-Medium (~100ms savings) |
| 5 | Strong name signing (for NGEN) | Low | Enables #1 |
| 6 | Callback caching | High | High (runtime performance) |
| 7 | Debug string optimization | Low | Low |
| 8 | Host detection streamlining | Low | Minimal |

---

## 🔍 Quick Wins (Config Changes Only)

These can be done immediately without code changes:

1. **Fix app.config binding redirects** - Update versions to `1.3.5.0`
2. **Add NGEN commands to installer** - See Section 1
3. **Set `use_keymouse_hooks=false`** in config.txt if hooks aren't needed
4. **Verify `async_startup=true`** is enabled (already set in current config)

---

## Testing Recommendations

After implementing optimizations:

1. **Measure baseline** before changes using Stopwatch in `OnConnection2`
2. **Test each optimization individually** to measure impact
3. **Monitor memory usage** - some optimizations trade memory for speed
4. **Test across Office versions** - Office 2010, 2013, 2016, 2019, 365
5. **Test 32-bit and 64-bit** Office installations

### Simple timing instrumentation:

```csharp
private Stopwatch startupTimer = new Stopwatch();

public void OnConnection2(object application)
{
    startupTimer.Start();
    // ... existing code ...
}

private void BootstrapAddIn()
{
    // ... existing code ...
    startupTimer.Stop();
    DebugMessage($"Total startup time: {startupTimer.ElapsedMilliseconds}ms");
}
```

---

## Related Documentation

- [OPTIMIZATION.md](../../OPTIMIZATION.md) - Python-side optimization strategies
- [CLAUDE.md](../../CLAUDE.md) - Architecture overview and development guide
