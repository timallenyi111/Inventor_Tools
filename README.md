## Architecture Overall Goals

- [] Move the copy and replace logic out of the Assembly Class and into a module in the main program
- [] Make a frame class that handles generating and replacing ids for frames.

### Importing Strategy
1.	Read "root assembly document" and create an assembly occurrence with necessary input data
2.	Create a part occurrence for each part in the assembly and add it to the assembly occurrence
3. Create a frame occurrence for every frame generator file and store it in the assembly 
4. Create an assembly occurrence for each assembly in the root assembly document and add it to the assembly occurrence
5. Repeat steps 2-4 for each assembly occurrence created in step 4

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


