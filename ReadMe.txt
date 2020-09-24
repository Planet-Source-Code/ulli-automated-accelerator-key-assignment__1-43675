Automated Accelerator Key Assignment

This Add-In analyses all (or selected) controls in a form and assigns or generates the necessary accelerator hotkeys, ie those keys that you press together with the Alt-key to access a control (the underlined characters which are preceeded by an ampersand in code). The problem is that a different accelerator should be selected from the caption of each individual control. To achieve this one has to generate all possible permutations to find the best solution where the accelerators are all different and as far to the left as possible.

When no control is selected in a form the add-in assumes that you want to process all controls. If you want to process selected controls only you must select them before running this add-in, either by enclosing them with the mouse or by clicking them with the Cntl key down.

Three algorithms for selecting the hotkeys are available:

Dumb 
----
just does one pass over all (or selected) controls and assigns hotkeys as necessary. If there are clashes they will not be resolved.

Smart
-----
makes as many passes as are necessary to arrive at a clash-less solution. If there is no such solution this is similar to the Optimum-algorithm because all possible permutations are examined, however a timer will stop the search after one minute.

Optimum (brute force)
-------
makes as many passes as are necessary to arive at the optimum clash-less solution. This will take very long because all possible permutations are examined in a brute force manner. Note that ten (10 only) controls will result in 10(factorial) equalling 3,628,800 permutations.

In a clash-less solution a different hotkey is assigned to each and every control. Only controls which have a caption property are examined.

Compile the DLL into your VB directory and then use the Add-Ins Manager to load the Hotkey Add-In into VB.
