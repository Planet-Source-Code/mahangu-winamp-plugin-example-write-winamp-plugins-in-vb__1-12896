THIS README FILE WAS THE ONE THAT CAME WITH THE TEMPLATE!!!!


ReadMe for the Template Folder
------------------------------

[Col_Rjl 27/09/2000 col_rjl@hotmail.com]

This folder contains the Visual Basic files for the template plugin.
This template is a minimal but functional GenWrapper style plugin. The
aim is that this can be used as a starting point for writing a proper
plugin. To create a new WinAmp General Purpose plugin using the
GenWrapper architecture from this template, the following steps should
be undertaken.

1) Copy this whole directory (..\Template) and give it a decent name,
eg the name you want to give your plugin. (It is probably best if
this Template folder and your new folder are in the same directory
so VB can find the type library, but this isn't too important.)

2) Go into the directory and rename Template.vbp to <ProjectName>.vbp
where ProjectName is the name you want your dll to have. You don't
have to do this, but it helps.

3) Each plugin written in VB using the GenWrapper architecture needs
its own copy of GenWrapper.dll. So copy GenWrapper.dll (from the root
of this distribution) into your directory. You have to rename it as
gen_<ProjectName>.<ClassName>.dll. ClassName will be "Plugin" by
default, so if your ProjectName was Smith, you would rename
GenWrapper.dll to gen_Smith.Plugin.dll.

4) Open the vbp file in Visual Basic. There is only one other file -
the Plugin class module. This contains more information about what
you have to write inside your plugin. So read this for further
instructions. The MOST IMPORTANT thing to do is to rename the
project: I'm not talking about the .vbp file, but the internal name
inside Visual Basic. You change this in the Project Explorer. It is
currently called GenTemplate. CHANGE THIS!

5) Once you have written your plugin in VB, you will have created a
dll called <ProjectName>.dll. When you distribute your plugin, both
<ProjectName>.dll and gen_<ProjectName>.<ClassName>.dll must be
copied into the WinAmp plugins directory. Once the VB dll is made,
it must never be renamed. It is crucial that the dll and the project
share the same name, and that it is this name that forms part of the
name given to GenWrapper.dll.

The VbSample shipped with this distribution demonstrates a plugin
that actually does something useful. The aim of the template is so
that there is something to copy whenever you want to make a new
plugin.

Enjoy.
