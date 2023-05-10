# CATIA Product Configurator

Welcome to the CATIA Product Configurator repo! This CATIA product configurator is based on the methods taught in the product modeling course at LiU ([TMKT57](https://studieinfo.liu.se/en/kurs/TMKT57/vt-2022)). It is developed for the [Design Automation Lab](https://liu.se/en/research/design-automation-lab) at the division of product realisation at Linköping University as part of my master thesis.

This is an example repo for how to implement a product configurator and STL exporter in CATIA. The STLs produced by this repo is intended to be utilized in the implementation of a product configurator in WebGL according to the [webgl-product-configurator repo](https://github.com/patrikdolsson/webgl-product-configurator). This repo can be used as a template to implement your own product configurator and STL exporter for any other customizable product. The product configurator is created using a windows forms application in Visual Studio connected to the [CATIA API](https://catiadesign.org/_doc/V5Automation/) using COM-referenences. 

A [startup guide](#getting-started), an [explanation of how the product configurator works](#how-does-it-work) and a [brief step by step guide to implement your own product configurator](#how-to-implement-your-own-product-configurator) will follow.

## Getting started

How to run the example:

-   Download/clone this repo to a folder on your machine, or import this repo to your own repo and clone that repo
-   Unzip the CATIA files and place it somewhere on your local machine
-   Open the project solution in Visual Studio
-   Open FileLocation.xml and replace the paths to corresponding desired locations on your local machine
-   Open TableLamp.CATProduct in CATIA and make sure you are in the assembly design workbench.
-   Run the solution and test the product configurator

## How does it work?


## How to implement your own product configurator