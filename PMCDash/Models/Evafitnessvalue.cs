namespace PMCDash.Models
{
    public class Evafitnessvalue
    {
        public int Idx { get; set; }
        public int Fitness { get; set; }

        public Evafitnessvalue(int Idx, int Fitness)
        {
            this.Idx = Idx;
            this.Fitness = Fitness;
        }
    }
}