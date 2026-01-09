### Base To-Do List
- [ ] Check if 2 different component files have the same name but different content
- [x] Figure out how to handle bolted connections and other content center objects
- [x] Make Frame Generator Members save in a sub-directory "\(parent file name\Frame\)"
- [x] Show on the form when the copy process is complete
- [ ] Make an installer

###To-Do before v0.1
- [x] Make a Thumbnail
- [x] Allow renaming of individual components
- [x] Remove the original tree node
- [x] Allow the option to not copy certain components
- [x] If an assembly is not being copied then all of its parts can't be copied either
- [x] Remove the main menu for v0.1
- [x] Modernize the UI

### Notes
Frame Attributes:
att name: Type
att value: MasterFrameOcc

### Issues
#### Frame replacement doesn't always work  
- ~~Current fix is replacing parts before assemblies, I don't know if that order is the actual issue or if it is
a timing issue with onedrive.~~
- Issue may be happening when if another assembly that contains the assembly being copied is open.  
	- Potential fix would be to check all open assemblies for the parts/assemblies being replaced and close them first.


