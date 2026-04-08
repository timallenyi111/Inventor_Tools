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
- [] There seems to be an issue with suppressed parts. 
- [] When trying to add duplicate parts to the highlight set, some items aren't added.
#### Frame replacement doesn't always work  
- ~~Current fix is replacing parts before assemblies, I don't know if that order is the actual issue or if it is
a timing issue with onedrive.~~
- Issue may be happening when if another assembly that contains the assembly being copied is open.  
	- Potential fix would be to check all open assemblies for the parts/assemblies being replaced and close them first.
- Another potential cause is when the frame is in a sub-assembly, not just a frame in the root assembly
	- To fix this we probably need to replace the skeleton frame id before saving the frame
	- Then open the parent assembly to the frame and replace it. 


