/******************************************************************************
 * iDempiere Plugin Blank Starter                                              *
 * Copyright (C) Contributors                                                  *
 *                                                                            *
 * This is a blank starter template. Rename the package and bundle to match   *
 * your module name (e.g., org.idempiere.mymodule).                           *
 *****************************************************************************/
package org.idempiere.blank;

import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;

/**
 * Plugin Activator - OSGI bundle entry point
 *
 * OPTIONS:
 *
 * 1. For plugins WITH 2Pack (AD models):
 *    - Change "implements BundleActivator" to "extends Incremental2PackActivator"
 *    - Remove start/stop methods (parent handles them)
 *    - Place 2Pack_x.x.x.zip files in META-INF/
 *
 * 2. For plugins WITHOUT 2Pack:
 *    - Use this simple activator as-is
 *    - Add your initialization code in start()
 */
public class Activator implements BundleActivator {

    private static BundleContext context;

    public static BundleContext getContext() {
        return context;
    }

    @Override
    public void start(BundleContext bundleContext) throws Exception {
        Activator.context = bundleContext;
        // Add your plugin startup code here
    }

    @Override
    public void stop(BundleContext bundleContext) throws Exception {
        // Add your plugin shutdown code here
        Activator.context = null;
    }
}
