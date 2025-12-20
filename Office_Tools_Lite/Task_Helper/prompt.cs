namespace Office_Tools_Lite.Task_Helper
{
    public class Prompt
    {
        public async Task<string> PrompterK()
        {
            var firstpart = "31U8SWmt3934lIVX6J5eYRFBTGT32m0bvG+JW4EsQ6U=";
            
            return await Task.FromResult(firstpart);
        }

        public async Task<string> PrompterI()
        {
            var secondpart = "GfylevKHWeAdcHoGsQfxiw==";
            return await Task.FromResult(secondpart);
        }

        
    }
}