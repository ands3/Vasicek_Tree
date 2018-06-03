# Vasicek_Tree

This Excel worksheet constructs a Vasicek interest rate binomial tree in VBA based on the Vasicek tree 
model construction procedure in Bruce Tuckman's Fixed Income Securities 2nd edition (page 232-244). 
Using a set of user given inputs, Problem2 (i) tab builds a Vasicek interest rate binomial tree by first 
computing the middle nodes. It then proceeds to build nodes in the upper half and lower half of the tree. 
During this process, it creates additional (secondary) nodes that are used for computational purposes 
only. After the initial build, it then clears any secondary nodes by simply changing their color to white.
