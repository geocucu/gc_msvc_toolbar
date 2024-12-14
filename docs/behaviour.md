#### Property persistence 

The persistence type (user or project) in the property sheet config XMLs is not a hard rule for what to specify with 
get/set property methods.  
A property may appear in both the user and project files at the same time. In that case, VS will pick up whichever
file matches the "canonical" persistence.

#### Property update

Changes in the project property sheet followed by "Apply/OK" will immediately be saved to disk, to the file 
corresponding to the project's canonical persistence ("user"/"project").  
Hence monitoring the files directly is a possible choice for monitoring changes to the project configuration. 
Risky, but the alternative (`IVs[...]Events`) interfaces make me not care right now. 

#### Startup Project

Changing it doesn't affect either the project user/project files, or the solution file.  
Most likely the hidden `.vs` dir contains some flag somewhere with the Startup Project. 
Deleting `.vs` resets the Startup Project selection.