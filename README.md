# CATIA Product Configurator

Welcome to the CATIA Product Configurator repo! This CATIA product configurator is based on the methods taught in the product modeling course at LiU ([TMKT57](https://studieinfo.liu.se/en/kurs/TMKT57/vt-2022)). It is developed for the [Design Automation Lab](https://liu.se/en/research/design-automation-lab) at the division of product realisation at Linköping University as part of my master thesis.

This is an example repo for how to implement a product configurator and STL exporter in CATIA. The STLs produced by this repo is intended to be utilized in the implementation of a product configurator in WebGL according to the [webgl-product-configurator repo](https://github.com/patrikdolsson/webgl-product-configurator). This repo can be used as a template to implement your own product configurator and STL exporter for any other customizable product. The product configurator is created using a windows forms application in Visual Studio connected to the [CATIA API](https://catiadesign.org/_doc/V5Automation/) using COM-referenences. 

A [startup guide](#getting-started), a [preview](#preview), an [explanation of how the product configurator works](#how-does-it-work) and a [brief step by step guide to implement your own product configurator](#how-to-implement-your-own-product-configurator) will follow.

## Getting started

How to run the example:

-   Download/clone this repo to a folder on your machine, or import this repo to your own repo and clone that repo
-   Unzip the CATIA files and place it somewhere on your local machine
-   Open the project solution in Visual Studio
-   Open FileLocation.xml and replace the paths to corresponding desired locations on your local machine (Note: The STL and temp paths do not have to exist)
-   Open TableLamp.CATProduct in CATIA and make sure you are in the assembly design workbench.
-   Run the solution and test the product configurator

## Preview

The windows forms app should look like the following:

![Product configurator](readme-images/product-configurator.png)

The resulting product in CATIA should look something like the following when configured:

![CAD model](readme-images/CAD%20model.png)

## How does it work?

The configurator uses abstracted [power copies](http://catiadoc.free.fr/online/pktug_C2/pktugat0053.htm) to instantiate parts in desired context. The entire product is defined in code by fully defining the instantiation of all the parts in terms of the configuration parameters. All context references are hard coded as defined by the power copies. 

The parameters that affect the geometry of components of the product will need to be discretized when exporting the STL models for use in the [webgl implementation](https://github.com/patrikdolsson/webgl-product-configurator). This discretization of parameters and determining of STL models that are exported using the export STL button is done by using what I call a PCM (Permutation Calculation Matrix) as shown in the following figure:

![Permutation Calculation Matrix](readme-images/PCM.png)

The setup of part of this matrix is done as self-evident in windows forms application as shown in the [preview](#preview). The totals on the bottom row is the number of configurations of each instantiated component as according to the discretization of the parameters. The numbers of configurations for each component are calculated using the following formula:

$$\text{Permutations of a component} = \prod_{i=1}^{n}\left(p_{i,s} + 1\right)$$

where,

-   $n$ = the number of geometry affecting parameters
-   $p_{i,s}$ = the number of uniform discretized steps that the parameter $p_i$ can take inside its interval

The PCM is derived from what I call a GAPASM (Geometry-Affecting Parameter Associative Structure Matrix), which is a matrix that intends to show which parameters affect the geometry of what components. An example of this is shown in the following figure:

![Geometry-Affecting Parameter Associative Structure Matrix](readme-images/GAPASM.png)

The actual configurations are calculated using recursive functions in order to calculate all the different configurations no matter how many parameters are involved. 

Furthermore, the configuration of the GUI control in the [webgl implementation](https://github.com/patrikdolsson/webgl-product-configurator) is almost fully determined by the stlInfo.json file produced by this repo. Which parameters that should be available for change in the webgl implementation is determined by the explicit setup of the STL export in the module [ExportSTL.vb](https://github.com/patrikdolsson/CATIA-product-configurator/blob/main/CATIA%20Product%20Configurator/CATIA%20Product%20Configurator/ExportSTL.vb).

## How to implement your own product configurator

<!--The current implementation does not support an arbitrary amount of instantiations of a certain power copy out of the box-->