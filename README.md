# The-genetic-algorithm-VBA-code
## Implementing a genetic algorithm to optimize rock cutting tools

## Macro program description

The block diagram of the macro program for optimizing the tool geometry parameters depending on the operating conditions and the mechanical drilling speed is shown in Fig. 1.
The macro program performs the following procedures:
1. Set the initial version of the tool geometry by setting the values of the parameters X1,. . . , Xn, as well as several options with tool geometry parameters, consisting of random numbers.
2. The construction of the tool mesh is performed by substituting the values of the geometric parameters X1 ... Xn.
3. Next, the model file system is prepared, where a new finite element mesh of the tool is combined with the main grid of the soil, and a contact surface and other model parameters are created.
 
![Fig](https://user-images.githubusercontent.com/50267432/57181450-1bc13800-6ebe-11e9-9a41-660001d68809.GIF)

Fig. 1.The block diagram of the macro program

4. After the model is calculated from the output files, the values of the target functional F are determined, as well as the parameters specified as constraints.
5. Then the genetic algorithm gets the value of the functional F and generates a new set of parameters X1,. . . , Xn, after which the step «2» is repeated.
6. The calculation will be completed after the specified number of cycle steps has passed.
