# HE2MATH v. 1.0

This is the first release which I am going to call "stable". I will keep the syntax back-compatible with this version for all the minor upgrades of version `1`.

This version **is not back-compatible with `0.1` and `0.2`**, all the HE2MATH function calls have to be rewritten when you upgrade from `0.1` or `0.2`.

The distribution script: [he2math_1.0.zip][he2math_1_0_zip].

# System Requirements

- **Operating System**: Windows XP; Windows 7; Windows 8. I have not tested with Windows 10 yet. Let me know if you have!
- **Microsoft Office**: MS Office 2007; MS Office 2010; MS Office 2013.
- **Mathcad**: The only supported version is Mathcad 15.
- **HEPAK**: any version supporting MS Excel function calls (`xla`-macro version).

# Installation

Download and unpack `HE2MATH` script file anywhere you can easily find it later (e.g. your `Documents` folder).

# How to Use HE2MATH

**IMPORTANT: this manual applies to a version `1.0`.**

Add the reference to HE2MATH Mathcad file in the worksheet:

![][mc_reference1]
![][mc_reference2_he]

As soon as the reference is added, `HE2MATH` functions can be called within the worksheet. The following are few simple examples describing the syntax of `HE2MATH`.

The first example case defines the density of helium at atmospheric pressure (`101325 Pa`) and ambient temperature (`20 C`).

![][mc_example_01_he]

- First two lines define the pressure and temperature of helium at the point of interest. Defined variables `p.1` and `t.1` could be then used as parameters of the `he2math_density` function.
- Third line calls `he2math_density` function which returns the density of helium at `101325 Pa` and `20 C`.

## Syntax

- The order of inputs is as follows: `input_1_code`, `input_1_value`, `input_2_code`, `input_2_value`.
- The first and third input parameters of the function correspond to the first and second input codes as per hepak manual. Input codes can be entered as numbers or special variables starting with `property_` can be used (e.g. `property_pressure` or shortened `prop_p`).
- The second and forth parameters are input values describing the state of the fluid at the point of interest.
- Syntax of `he2math` scripts is similar to syntax of HEPAK for MS Excel. There are two main differences: the unit system (it's always SI, see the **Units** section below for detailed explanation) and the phase calculation code (since `0` covers all the scenarios, it's hard-coded into the script).

## Input Codes

- Input codes describe the properties to be entered as the second and forth parameters. In the example above, `property_pressure` input code indicates that the second parameter is pressure, `property_temperature` input code indicates that the forth parameter is temperature.
- Input codes can be entered as numbers, `property_pressure` and `property_temperature` are usual Mathcad variables containing values `1` and `2`.
- The list of possible input codes is defined by `he2math_properties` variable and can be displayed using equal sign anywhere in the worksheet.

![][mc_inputs_he]

## Units

- All the fluid properties (pressure, temperature, density, enthalpy, etc) inside of `HE2MATH` functions can be defined with units. Unitless variables are assumed to be in SI system.
- HE2MATH **does not** verify the validity of units, i.e. replacement of `Pa` with `kg` would lead to **the exact same result**. That is because `HE2MATH` only rescales all the units to their corresponding SI representations. Since `kg` is a standard SI unit, it is not rescaled after `HE2MATH` function call.
- All the values are returned by `HE2MATH` functions  with units. Returned values can be rescaled in a standard Mathcad-way by placing a unit next to the variable.

![][mc_example_02_he]

## Functions

- The name of the function defines the property to be returned.
- All the functions start with `he2math_` and are followed by the name of the property to be calculated.
- The list of the valid function names is defined by `he2math_functions` variable and can be displayed using equal sign anywhere in the worksheet.

![][mc_functions_he]

## Saturated conditions

- Calculations of fluid properties in case if they are fully defined by one variable (i.e. saturated temperature at a given pressure) could have been done using special input codes (e.g. `property_lambda_line` or `property_saturated_liquid`) followed by any value. HEPAK constants (accessible using `heconst()` function in Excel) are not yet available.
- The example below shows the calculation of temperature for saturated liquid helium at atmospheric pressure.

![][mc_example_03_he]


# HEPAK installation

In order to use `HE2MATH` functions within Mathcad worksheet, it is necessary to have HEPAK macro installed in MS Excel. If HEPAK is already installed (i.e. something like `=hecalc("d",0,"p",101325,"T",300,1)` in excel spreadsheet returns a valid result), ignore this part.

## Installation Process

HEPAK installation includes the only step of copying HEPAK distribution file (it is usually named `HePak.xla`) to `XLSTART` folder. This allows to load HEPAK plug-in automatically every time Excel is started.

Open Microsoft Office directory (usually `C:\Program Files\Microsoft Office` for Windows 32-bit version and `C:\Program Files (x86)\Microsoft Office` for Windows 64-bit version).

Depending from the version of MS Office installed, open the corresponding folder:

| **Office Version**   | **Folder Path**    |
|----------------------|--------------------|
| MS Office 2007 &nbsp;| `Office12\XLSTART` |
| MS Office 2010 &nbsp;| `Office14\XLSTART` |
| MS Office 2013 &nbsp;| `Office15\XLSTART` |

Copy `HePak.xla` file received from HEPAK distributor there.

![][he_inst1] 

To verify that you are done with HEPAK installation, open a new Excel document and type `=hecalc("d",0,"p",101325,"T",300,1)` in any cell (or just copy-paste that).

![][he_verif1]

If everything is OK you should see `0.162522`. The value shows density of helium in `kg/m^3` at `101325Pa` pressure and `300K` temperature.

![][he_verif2]    

[he_inst1]: ./images/he_inst1.png
[he_verif1]: ./images/he_verif1.png
[he_verif2]: ./images/he_verif2.png
[mc_example_01_he]: ./images/mc_example_01_he.png
[mc_example_02_he]: ./images/mc_example_02_he.png
[mc_example_03_he]: ./images/mc_example_03_he.png
[mc_functions_he]: ./images/mc_functions_he.png
[mc_inputs_he]: ./images/mc_inputs_he.png
[mc_reference1]: ./images/mc_reference1.png
[mc_reference2_he]: ./images/mc_reference2_he.png
[he2math_1_0_zip]: ../../raw/master/he2math_v1.0.zip
