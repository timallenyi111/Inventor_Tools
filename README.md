## Architecture To-do

- [x] Move the copy and replace logic out of the Assembly Class and into a module in the main program
- [x] Make a frame class that handles generating and replacing ids for frames.
- [] Make a bolted connection class that handles the file path logic
- [] Try to get frame copying to happen without having to press buttons in Inventor
	- a good place to start would be to see if we actually need to open the frame assembly (not the frame parent assembly) to replace it's component

### Importing Strategy
1.	Read "root assembly document" and create an assembly occurrence with necessary input data
2.	Create a part occurrence for each part in the assembly and add it to the assembly occurrence
3. Create a frame occurrence for every frame generator file and store it in the assembly 
4. Create an assembly occurrence for each assembly in the root assembly document and add it to the assembly occurrence
5. Repeat steps 2-4 for each assembly occurrence created in step 4

### Copy and Replace Strategy

1. Copy all part files
2. Copy all assemblies (not frames)
3. Use the original frame assembly document to replace the skeleton part, skeleton part id, and frame attribute skeleton id, then save a copy
4. Move all the way down the assembly tree to the lowest assembly and start replacing components working our way back up.




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


