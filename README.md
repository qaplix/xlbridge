# XlBridge
***Write user-defined Excel formulas in .NET (C#) or Python. No more VBA!***

## Excel formulas in C# or Python

XlBridge enables you to create Excel formulas using C# or Python using your standard toolset. Only a few lines of code is needed to expose an existing library to Excel.

Your custom code runs outside the Excel process, preventing your own bugs from crashing Excel and letting you update or add formulas without restarting Excel.

The XlBridge add-in communicates with your custom code through (local or remote) gRPC connections.

## Getting started

The essential code needed to expose C# functions to Excel is:

```csharp
    var testServer = new BridgeServiceBuilder()
        .AddFunctions.FromType<MyFunctions>()
        .CreateNativeGrpcServer();
    
    testServer.Start();
    
    await testServer.ShutdownAsync();
```

Read more at <https://xlbridge.qaplix.se>

## Sample repo

This repository contains sample code to use with the ***XlBridge Excel add-in*** and the *XlBridge user libraries*.

The issue tracker can be used for feedback or email [info@qaplix.se](mailto:info@qaplix.se)

## Copyright

Copyright © 2021 Qaplix AB

XlBridge is commercial software in public beta. Trial licenses and a free tier is available.

See the [xlbridge webpage](https://xlbridge.qaplix.se) for more information.
