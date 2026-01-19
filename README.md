# iDempiere Plugin Blank Starter

A minimal template for creating new iDempiere plugins.

## Quick Start

1. **Clone/Download this repo**
2. **Rename to your module** (e.g., `org.idempiere.mymodule`)
3. **Update these files:**
   - `META-INF/MANIFEST.MF` - Change Bundle-SymbolicName, Bundle-Name
   - `.project` - Change project name
   - `src/org/idempiere/blank/` - Rename package folder
   - `Activator.java` - Update package declaration

4. **Import into Eclipse** as an existing project
5. **Start coding!**

## Project Structure

```
org.idempiere.blank/
├── META-INF/
│   └── MANIFEST.MF      # OSGI bundle configuration
├── src/
│   └── org/idempiere/blank/
│       └── Activator.java   # Plugin entry point
├── plugin.xml           # Extension points (processes, callouts, etc.)
├── build.properties
├── .classpath
├── .project
└── README.md
```

## Adding 2Pack (AD Models)

If your plugin includes Application Dictionary models (tables, windows, etc.):

1. Change Activator to extend `Incremental2PackActivator`:
   ```java
   import org.adempiere.plugin.utils.Incremental2PackActivator;

   public class Activator extends Incremental2PackActivator {
       // No start/stop methods needed - parent handles 2Pack import
   }
   ```

2. Update MANIFEST.MF to include the plugin utils:
   ```
   Require-Bundle: org.adempiere.base;bundle-version="11.0.0",
    org.adempiere.plugin.utils;bundle-version="11.0.0"
   ```

3. Place your `2Pack_x.x.x.zip` files in `META-INF/`

## Using with Ninja

Generate this structure automatically from Excel:

```bash
# In org.idempiere.ninja directory
./RUN_Ninja.sh MyModule.xls -d
```

See: https://github.com/red1oon/org.idempiere.ninja

## Common Extension Points

Add to `plugin.xml`:

### Process
```xml
<extension point="org.adempiere.base.Process" id="myprocess">
   <process id="MyProcess" class="org.idempiere.mymodule.process.MyProcess"/>
</extension>
```

### Model Factory
```xml
<extension point="org.adempiere.base.ModelFactory" id="mymodel">
   <modelfactory class="org.idempiere.mymodule.model.MyModelFactory"/>
</extension>
```

### Callout
```xml
<extension point="org.adempiere.base.Callout" id="mycallout">
   <callout class="org.idempiere.mymodule.callout.MyCallout"/>
</extension>
```

### Event Handler
```xml
<extension point="org.adempiere.base.ModelValidator" id="myvalidator">
   <modelvalidator class="org.idempiere.mymodule.event.MyEventHandler"/>
</extension>
```

## Requirements

- iDempiere 11+
- Java 17+
- Eclipse IDE with PDE (Plugin Development Environment)

## License

GPL v2

---
*Created with Ninja - https://github.com/red1oon/org.idempiere.ninja*
